# START_MODULE_CONTRACT
# PURPOSE: Generic-движок аудита для всех типов с текстовым содержимым (docx/pptx/pdf) через multi_rule LLM-путь. Используется, когда в config.json типа задан `parser_by_ext` и нет поля `engine` (спецдвижки по xlsx — в реестре `src/engines.py`, раннеры в `src/doc_type_validators/`).
# INPUTS: doc_type (имя папки в doc_configs/), target file. Конфиг через `load_audit_config`. LLM через `run_multi_rule_audit`.
# OUTPUTS: `AuditResult` с violations, warnings и unchecked_rules + Excel-отчёт + JSON-логи в session_dir. Пустой/нераспознанный документ — ValueError до вызова модели (`check_document_text`). Непроверенные правила (модель не вернула вердикт) помечаются «НЕ ПРОВЕРЕНО» (F15); 0 проверенных — ValueError.
# KEYWORDS: engine, multi-rule, generic-runtime, docx, pptx, pdf.
# LINKS: src/audit/models.py (AuditConfig/AuditResult), src/audit/logger.py (PipelineLogger), src/audit/excel_reporter.py (save_to_excel), src/llm/multi_rule.py (run_multi_rule_audit), src/format_parsers/ (parse_docx/parse_pptx/parse_pdf), config/parsers.py (PARSERS_CONFIG).
# RATIONALE:
#   Раньше класс AuditEngine жил в main.py и был «исправлен» через 3 monkey-patch
#   функции (`_audit_engine_init/parse_secondary/list_doc_types`) в RUNTIME_INTEGRATION
#   из-за артефактов inline-монолита. Здесь класс — нормальный Python с правильными
#   путями (PROJECT_ROOT через parents[2]) и импортами. Никаких monkey-patch.
# END_MODULE_CONTRACT

from __future__ import annotations

# START_IMPORTS
import json
import re
import time
from datetime import datetime
from pathlib import Path
from typing import Any, Callable, Dict, List, Optional

from config.parsers import PARSERS_CONFIG
from src.audit.excel_reporter import save_to_excel
from src.audit.logger import PipelineLogger
from src.audit.models import AuditConfig, AuditResult, load_audit_config
from src.doc_type_parsers.grafik_obhod import parse_grafik_obhod
from src.format_parsers import parse_docx, parse_pptx
from src.format_parsers.pdf import parse_pdf
from src.format_parsers.pdf._clients import VLMClient
from src.llm import load_methodology_config, load_multi_rule_config, run_multi_rule_audit
from src.llm.multi_rule import build_prompts
# END_IMPORTS


# START_PATHS
# PURPOSE: Локальные пути модуля. parents[2] = repo root (engine.py лежит в src/audit/).
_PROJECT_ROOT = Path(__file__).resolve().parents[2]
_DOC_CONFIGS_DIR = _PROJECT_ROOT / "doc_configs"
# END_PATHS


# START_TEXT_GUARD
# PURPOSE: Проверка, что парсер извлёк из документа текст, прежде чем звать модель. Пустой документ
# (blank.docx, pptx из одних картинок, PDF с пустыми страницами) или документ из одних маркеров
# модель «проверяет» и выдумывает нарушения — аудит устойчивости, находки 1.5, 1.5b, 2.4f, 3.1b.
# Маркеры, которые текстом документа не считаются: границы страниц/слайдов и ошибки распознавания.
_STRUCTURAL_MARKER_RE = re.compile(r"^\[(СТРАНИЦА|СЛАЙД) \d+\]$")
_ERROR_MARKER_PREFIX = "[ОШИБКА"
# Минимум полезных символов (без маркеров), чтобы документ считался прочитанным. 1 = «хоть что-то»:
# больший порог мог бы отбраковать короткие настоящие документы.
_MIN_DOC_TEXT_CHARS = 1


def check_document_text(raw_text: str) -> None:
    """
    Назначение:
        Убедиться, что в `raw_text` есть текст документа, а не только маркеры.

    Вход:
        raw_text: полный текст документа от формат-парсера.

    Выход:
        None. Если полезного текста меньше `_MIN_DOC_TEXT_CHARS` — ValueError с человекочитаемой
        причиной (с текстом первой ошибки распознавания, если она была).

    Логика:
        Строки-маркеры границ отбрасываются, строки-маркеры ошибок собираются отдельно,
        остальное считается текстом документа.
    """
    useful_chars = 0
    errors: List[str] = []
    for line in raw_text.splitlines():
        s = line.strip()
        if not s or _STRUCTURAL_MARKER_RE.match(s):
            continue
        if s.startswith(_ERROR_MARKER_PREFIX):
            errors.append(s)
            continue
        useful_chars += len(s)
    if useful_chars >= _MIN_DOC_TEXT_CHARS:
        return
    if errors:
        raise ValueError(f"Документ не удалось распознать, проверка не выполнена. {errors[0]}")
    raise ValueError(
        "Документ не содержит распознаваемого текста (пустой файл или только изображения), проверка не выполнена."
    )
# END_TEXT_GUARD


# START_SECONDARY_PARSER_DISPATCH
# PURPOSE: Парсеры для secondary-файлов (multi-file аудиты типа DOCX + XLSX).
# Используется только внутри `AuditEngine._parse_secondary`.
SECONDARY_PARSER_DISPATCH: Dict[str, Callable[[str], Dict[str, str]]] = {
    "grafik_obhod": parse_grafik_obhod,
}


def get_parser(name: str) -> Callable[[str], Dict[str, str]]:
    """Возвращает зарегистрированный парсер вторичного файла по имени."""
    if name not in SECONDARY_PARSER_DISPATCH:
        raise KeyError(f"Парсер '{name}' не зарегистрирован. Доступные: {sorted(SECONDARY_PARSER_DISPATCH)}")
    return SECONDARY_PARSER_DISPATCH[name]
# END_SECONDARY_PARSER_DISPATCH


# START_DOC_TYPE_VALIDATION
# PURPOSE: Проверка каталога типа при листинге — чтобы повреждённый тип не выглядел рабочим и
# не падал пустым 500 в середине запуска. Проверяется ровно то, без чего аудит не стартует:
# config.json парсится, задан ровно один источник правды (parser_by_ext XOR engine), у generic-типа
# есть и парсятся sections.json и rules_multi.json.
def _json_error_reason(path: Path, exc: json.JSONDecodeError) -> str:
    """Человекочитаемая причина для битого JSON: имя файла и строка."""
    return f"{path.name} повреждён (строка {exc.lineno}, позиция {exc.colno})"


def _validate_doc_type_dir(config_dir: Path, entry: Dict[str, Any]) -> Optional[str]:
    """
    Назначение:
        Проверяет конфигурацию одного типа и заполняет `entry` (doc_type, doc_title).

    Вход:
        config_dir: каталог doc_configs/<тип>.
        entry: словарь записи списка типов — дополняется реальными doc_type/doc_title,
            если config.json читается.

    Выход:
        None, если тип рабочий; иначе строка причины (для `broken_reason`).
    """
    config_file = config_dir / "config.json"
    try:
        with open(config_file, "r", encoding="utf-8") as f:
            data = json.load(f)
    except json.JSONDecodeError as e:
        return _json_error_reason(config_file, e)
    except OSError as e:
        return f"config.json не читается: {e.strerror or e}"
    if not isinstance(data, dict):
        return "config.json: ожидался объект"
    entry["doc_type"] = data.get("doc_type", config_dir.name)
    entry["doc_title"] = data.get("doc_title", "")
    parser_by_ext = data.get("parser_by_ext", {})
    engine = data.get("engine")
    has_parser_map = isinstance(parser_by_ext, dict) and bool(parser_by_ext)
    has_engine = isinstance(engine, str) and bool(engine)
    if has_parser_map and has_engine:
        return "config.json: одновременно заданы parser_by_ext и engine"
    if not has_parser_map and not has_engine:
        return "config.json: не задан ни parser_by_ext (generic), ни engine (special)"
    if has_engine:
        return None
    # generic-LLM тип: без sections.json + rules_multi.json multi-rule аудит не стартует
    for name, key in (("sections.json", "sections"), ("rules_multi.json", "rules")):
        path = config_dir / name
        if not path.exists():
            return f"нет файла {name}"
        try:
            content = json.loads(path.read_text(encoding="utf-8"))
        except json.JSONDecodeError as e:
            return _json_error_reason(path, e)
        except OSError as e:
            return f"{name} не читается: {e.strerror or e}"
        if not isinstance(content, dict) or key not in content:
            return f"{name}: нет ключа «{key}»"
    return None
# END_DOC_TYPE_VALIDATION


# START_AUDIT_ENGINE
class AuditEngine:
    """
    Единый generic-движок аудита документов через multi_rule LLM-путь.

    Использование:
        engine = AuditEngine("presentation_eu")
        result = engine.run("document.pptx")

    Или с кастомным base_dir:
        engine = AuditEngine("prikaz_ic", base_dir="/path/to/project")

    Special-движки (kpsc, kartochka_proekta, drivers, plan_grafik) НЕ используют
    этот класс — у них свои runner-ы в `src/doc_type_validators/`.
    """

    def __init__(
        self,
        doc_type: str,
        base_dir: Optional[str] = None,
        config_dir: Optional[str] = None,
    ):
        """
        Инициализация движка.

        Args:
            doc_type: тип документа (имя папки в doc_configs/).
            base_dir: корневая директория проекта (по умолчанию — корень репо).
            config_dir: путь к папке конфигов (переопределяет base_dir + doc_configs).
        """
        self.doc_type = doc_type
        self.base_dir = Path(base_dir) if base_dir else _PROJECT_ROOT
        self.config_path = Path(config_dir) if config_dir else self.base_dir / "doc_configs" / doc_type
        self.config: AuditConfig = load_audit_config(self.config_path)
        # Парсер-конфиг — singleton PARSERS_CONFIG из config/parsers.py.
        # Берём ссылки на нужные суб-конфиги, чтобы не тянуть в runtime сам объект.
        self.pdf_parser_config = PARSERS_CONFIG.pdf
        self.docx_header_ocr_config = PARSERS_CONFIG.docx_header_ocr
        self.logger: Optional[PipelineLogger] = None

    def run(
        self,
        target_path: str,
        model: Optional[str] = None,
        temperature: Optional[float] = None,
        parse_only: bool = False,
        print_prompts: bool = False,
        session_dir: Optional[str] = None,
        out_xlsx: Optional[str] = None,
        secondary_path: Optional[str] = None,
        progress_callback: Optional[Callable] = None,
    ) -> AuditResult:
        """
        Запуск полного цикла аудита через multi_rule LLM-путь.

        Args:
            target_path: путь к целевому документу.
            model: модель LLM (переопределяет конфиг).
            temperature: температура (переопределяет конфиг).
            parse_only: режим только парсинга.
            print_prompts: режим отладки — напечатать промпты обоих слоёв (base, methodology)
                и выйти без вызова модели (F18; раньше флаг принимался, но аудит шёл через модель).
            session_dir: директория для логов.
            out_xlsx: путь для сохранения Excel.
            secondary_path: путь к вторичному файлу (XLSX для multi-file аудитов).
            progress_callback: колбэк прогресса для веб-UI (type: str, data: dict).

        Returns:
            AuditResult с нарушениями и статистикой.
        """
        start_time = time.time()

        def _emit(event_type: str, data: dict = None):
            if progress_callback:
                progress_callback(event_type, data or {})

        model = model or self.config.model
        temperature = temperature if temperature is not None else self.config.temperature
        if session_dir:
            session_path = Path(session_dir)
        else:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            session_path = self.base_dir / "logs_result" / self.doc_type / f"session_{timestamp}"
        self.logger = PipelineLogger(session_path, self.config.doc_title)
        self.logger.log(f"🚀 Запуск аудита: {self.config.doc_title} ({self.doc_type})")
        self.logger.log(f"   Сессия: {session_path}")
        self.logger.log(f"   Целевой документ: {target_path}")
        if secondary_path:
            self.logger.log(f"   Вторичный файл: {secondary_path}")
        self.logger.log(f"   Модель: {model}")
        # Логируем карту парсеров по расширениям; блоки дополнительного контекста
        # выводим, только если соответствующий парсер реально задействован хотя бы для
        # одного расширения.
        self.logger.log(f"   Парсеры по расширениям: {self.config.parser_by_ext}")
        parser_names = set(self.config.parser_by_ext.values())
        if "paddle" in parser_names:
            # Значения берём из PDF infra-конфига (config/parsers.json), не из AuditConfig:
            # AuditConfig описывает doc_type, а сервисы PDF общие для всего рантайма.
            pdf_cfg = self.pdf_parser_config
            layout_src = pdf_cfg.layout.base_url or f"local: {pdf_cfg.layout.model}"
            self.logger.log(f"   Layout: {layout_src}")
            self.logger.log(f"   VLM: {pdf_cfg.vlm.model or 'авто (через /v1/models)'}")
            self.logger.log(f"   VLM URL: {pdf_cfg.vlm.base_url}")
        if self.config.llm_base_url:
            self.logger.log(f"   LLM URL: {self.config.llm_base_url}")
        if not Path(target_path).exists():
            self.logger.log(f"❌ Целевой документ не найден: {target_path}")
            raise FileNotFoundError(f"Целевой документ не найден: {target_path}")
        if self.config.secondary_file and (not secondary_path):
            self.logger.log(f"⚠️ Конфиг требует вторичный файл ({self.config.secondary_file.type}), но --secondary не указан")
        if secondary_path and (not Path(secondary_path).exists()):
            self.logger.log(f"❌ Вторичный файл не найден: {secondary_path}")
            raise FileNotFoundError(f"Вторичный файл не найден: {secondary_path}")
        # session_dir в событии — чтобы веб-бэкенд знал каталог сессии и при ошибке (F19).
        _emit("audit_start", {"doc_type": self.doc_type, "filename": Path(target_path).name, "session_dir": str(session_path)})
        target_ext = Path(target_path).suffix.lower()
        if self.config.secondary_file and target_ext == f".{self.config.secondary_file.type}" and (not secondary_path):
            self.logger.log(f"📊 Загружен {target_ext} — парсим как вторичный файл")
            _emit("parsing_target", {})
            target_doc = self._parse_secondary(target_path)
            _emit("parsing_target_done", {})
        else:
            _emit("parsing_target", {})
            target_doc = self._parse_document(target_path, session_path / "parse_logs")
            _emit("parsing_target_done", {})
            if secondary_path and self.config.secondary_file:
                secondary_chunks = self._parse_secondary(secondary_path)
                target_doc.update(secondary_chunks)
                self.logger.log(f"   📊 Merged {len(secondary_chunks)} чанков из вторичного файла")
        if parse_only:
            self.logger.log("✅ Режим --parse-only: парсинг завершён")
            print(f"\n{'=' * 60}")
            print("TARGET:")
            print(json.dumps(target_doc, ensure_ascii=False, indent=2))
            return AuditResult(doc_type=self.doc_type, session_dir=session_path, target_path=target_path, duration_sec=time.time() - start_time)
        # Guard: пустой/нераспознанный документ к модели не идёт. Путь «xlsx как вторичный файл»
        # отдаёт чанки без raw_text — его guard не касается.
        raw_text = target_doc.get("raw_text")
        if isinstance(raw_text, str):
            try:
                check_document_text(raw_text)
            except ValueError as e:
                self.logger.log(f"❌ {e}")
                raise
        # Предупреждения парсера (например, не распознанные страницы PDF) — в лог, результат и Excel.
        warnings: List[str] = list(target_doc.get("warnings") or [])
        for w in warnings:
            self.logger.log(f"⚠️ {w}")
        # Активный LLM-путь — multi_rule: парсер отдаёт raw_text, промпт собирается
        # по sections.json + rules_multi.json / rules_methodology.json.
        doc_configs_dir = self.config.config_dir.parent
        mr_config = load_multi_rule_config(doc_configs_dir, self.doc_type)
        if mr_config is None:
            raise ValueError(
                f"Для doc_type={self.doc_type} не найдены sections.json + rules_multi.json. "
                f"Без них multi-rule аудит запустить нельзя; legacy single-rule путь удалён."
            )
        # F4: предупреждения загрузчика правил (битый rules_custom.json) — к предупреждениям парсера.
        for w in mr_config.get("warnings", []):
            warnings.append(w)
            self.logger.log(f"⚠️ {w}")
        if print_prompts:
            # F18: показать, что увидит модель, по обоим слоям — и выйти, модель не вызывать.
            meth_pp = load_methodology_config(doc_configs_dir, self.doc_type)
            layers = [("base", mr_config["rules"], mr_config.get("include_scopes"))]
            if meth_pp is not None:
                layers.append(("methodology", meth_pp["rules"], meth_pp.get("include_scopes") or mr_config.get("include_scopes")))
            for layer, layer_rules, scopes in layers:
                system_prompt, user_prompt = build_prompts(
                    parsed=target_doc, sections=mr_config["sections"], rules=layer_rules,
                    include_scopes=scopes, filename=target_doc.get("filename", Path(target_path).name),
                )
                print(f"\n{'=' * 60}\n{layer.upper()} — SYSTEM PROMPT ({len(layer_rules)} правил)\n{'=' * 60}\n{system_prompt}")
                print(f"\n{'=' * 60}\n{layer.upper()} — USER PROMPT\n{'=' * 60}\n{user_prompt}")
            self.logger.log(f"✅ Режим --print-prompts: промпты {len(layers)} слоёв напечатаны, модель не вызывалась")
            return AuditResult(doc_type=self.doc_type, session_dir=session_path, target_path=target_path, duration_sec=time.time() - start_time)
        self.logger.log(f"🔍 Запуск multi-rule аудита (базовый слой, {len(mr_config['rules'])} правил)...")
        _emit("checking_rules", {"total": len(mr_config["rules"])})
        mr_result = run_multi_rule_audit(
            parsed=target_doc,
            sections=mr_config["sections"],
            rules=mr_config["rules"],
            include_scopes=mr_config.get("include_scopes"),
            filename=target_doc.get("filename", Path(target_path).name),
            llm_base_url=self.config.llm_base_url,
            llm_model=self.config.model,
            session_dir=session_path,
            layer="base",
            progress_callback=_emit,
        )
        violations = mr_result["violations"]
        # F15: правила без вердикта — по слоям, чтобы пометить «НЕ ПРОВЕРЕНО», а не «пройдено».
        unchecked_rules: List[Dict[str, Any]] = [
            {"index": r.get("index"), "title": r.get("title", ""), "layer": "base"}
            for r in mr_result.get("unchecked", [])
        ]
        self.logger.log(f"   Base usage: prompt={mr_result['usage']['prompt_tokens']} completion={mr_result['usage']['completion_tokens']} total={mr_result['usage']['total_tokens']}")
        meth_config = load_methodology_config(doc_configs_dir, self.doc_type)
        if meth_config is not None:
            self.logger.log(f"🔍 Запуск методического слоя ({len(meth_config['rules'])} правил, источник: {meth_config.get('source', 'МР/МУ')[:80]})...")
            meth_result = run_multi_rule_audit(
                parsed=target_doc,
                sections=mr_config["sections"],
                rules=meth_config["rules"],
                include_scopes=meth_config.get("include_scopes") or mr_config.get("include_scopes"),
                filename=target_doc.get("filename", Path(target_path).name),
                llm_base_url=self.config.llm_base_url,
                llm_model=self.config.model,
                session_dir=session_path,
                layer="methodology",
                progress_callback=_emit,
            )
            violations.extend(meth_result["violations"])
            unchecked_rules.extend(
                {"index": r.get("index"), "title": r.get("title", ""), "layer": "methodology"}
                for r in meth_result.get("unchecked", [])
            )
            self.logger.log(f"   Methodology usage: prompt={meth_result['usage']['prompt_tokens']} completion={meth_result['usage']['completion_tokens']} total={meth_result['usage']['total_tokens']}")
            self.logger.log(f"   Methodology violations: {len(meth_result['violations'])}")
        _all_multi_rules = [{**r, "layer": "base"} for r in mr_config["rules"]]
        if meth_config is not None:
            _all_multi_rules.extend(({**r, "layer": "methodology"} for r in meth_config["rules"]))
        # F15: фактически проверено = все правила минус непроверенные. Если 0 — ошибка,
        # а не «успешный» аудит с пустым/выдуманным результатом.
        checked_count = len(_all_multi_rules) - len(unchecked_rules)
        if checked_count <= 0:
            self.logger.log("❌ Модель не вернула ни одного вердикта — проверка не выполнена")
            raise ValueError("Модель не вернула ни одного вердикта — проверка не выполнена.")
        if unchecked_rules:
            self.logger.log(f"⚠️ Не проверено правил: {len(unchecked_rules)} из {len(_all_multi_rules)} (помечены «НЕ ПРОВЕРЕНО»)")
        _emit("checking_rules_done", {"violations": len(violations)})
        self.logger.log_final_results(violations)
        print(json.dumps(violations, ensure_ascii=False, indent=2))
        if out_xlsx:
            xlsx_path = out_xlsx
        else:
            xlsx_path = str(session_path / "audit_result.xlsx")
        save_to_excel(violations, xlsx_path, multi_rules=_all_multi_rules, warnings=warnings, unchecked=unchecked_rules)
        self.logger.log(f"📊 Excel сохранён: {xlsx_path}")
        duration = time.time() - start_time
        self.logger.log(f"{'=' * 60}")
        self.logger.log(f"📊 Итого нарушений: {len(violations)}")
        if warnings:
            self.logger.log(f"⚠️ Предупреждений парсера: {len(warnings)}")
        if violations:
            by_rule: Dict[Any, int] = {}
            for v in violations:
                idx = v.get("rule_index", "?")
                by_rule.setdefault(idx, 0)
                by_rule[idx] += 1
            self.logger.log("   По правилам:")
            for idx in sorted(by_rule.keys()):
                self.logger.log(f"   - Правило #{idx}: {by_rule[idx]} нарушений")
        self.logger.log(f"✅ Аудит завершён за {duration:.1f} сек. Сессия: {session_path}")
        return AuditResult(
            violations=violations, doc_type=self.doc_type, session_dir=session_path,
            duration_sec=duration, rules_checked=checked_count, target_path=target_path,
            warnings=warnings, unchecked_rules=unchecked_rules,
        )

    def _parse_document(self, file_path: str, parse_log_dir: Path) -> Dict[str, Any]:
        """
        Парсит документ выбранным парсером.

        Выбор парсера: `config.parser_by_ext[<расширение файла>]`. Никаких
        автоматических переопределений — один конфиг описывает всю цепочку.
        Если расширение не описано в карте, бросается ValueError.
        """
        file_ext = Path(file_path).suffix.lower()
        # Единый источник правды — parser_by_ext. Никаких скрытых override по формату.
        if file_ext not in self.config.parser_by_ext:
            raise ValueError(f"Для doc_type={self.doc_type} не сконфигурирован парсер для расширения {file_ext!r}. parser_by_ext={self.config.parser_by_ext}")
        parser_name = self.config.parser_by_ext[file_ext]
        if parser_name == "docx":
            self.logger.log(f"📄 DOCX-парсинг: {file_path}...")
            # Один и тот же VLMClient используется и для PDF-страниц, и для OCR шапок DOCX.
            vlm_cfg = self.pdf_parser_config.vlm
            vlm_client = VLMClient(
                base_url=vlm_cfg.base_url,
                model=vlm_cfg.model,
                api_key=vlm_cfg.api_key,
                max_tokens=vlm_cfg.max_tokens,
                temperature=vlm_cfg.temperature,
            )
            doc = parse_docx(file_path, vlm=vlm_client, header_ocr_cfg=self.docx_header_ocr_config)
        elif parser_name == "pptx":
            self.logger.log(f"📄 PPTX-парсинг: {file_path}...")
            doc = parse_pptx(file_path)
        elif parser_name == "paddle":
            self.logger.log(f"📄 PDF-парсинг (Paddle): {file_path}...")
            doc = parse_pdf(file_path, self.pdf_parser_config, log_dir=parse_log_dir)
        else:
            # Единственные поддерживаемые парсеры сейчас — docx/pptx/paddle. Любое другое
            # значение — ошибка конфигурации и причина явно падать, а не молча продолжать.
            raise ValueError(f"Неподдерживаемый парсер {parser_name!r} для {file_ext!r}. Допустимы: docx, pptx, paddle.")
        self.logger.log_parsed_doc(doc, Path(file_path).stem)
        return doc

    def _parse_secondary(self, file_path: str) -> Dict[str, Any]:
        """
        Парсит вторичный файл через зарегистрированный парсер
        (`SECONDARY_PARSER_DISPATCH` на уровне модуля).

        Добавляет префикс из config.secondary_file.chunk_prefix к ключам,
        а также ключ <prefix>filename с именем файла.
        """
        spec = self.config.secondary_file
        parser_fn = get_parser(spec.parser)
        raw_chunks = parser_fn(file_path)
        prefixed = {f"{spec.chunk_prefix}{k}": v for k, v in raw_chunks.items()}
        prefixed[f"{spec.chunk_prefix}filename"] = Path(file_path).name
        self.logger.log_parsed_doc(prefixed, f"secondary_{spec.type}")
        return prefixed

    @staticmethod
    def list_doc_types(base_dir: Optional[str] = None) -> List[Dict[str, Any]]:
        """
        Список доступных типов документов.

        Сканирует doc_configs/ и возвращает список {doc_type, doc_title}. Тип с повреждённой
        конфигурацией остаётся в списке, но помечается `broken: True` и `broken_reason`
        (текст для администратора) — интерфейс показывает его серым, запуск аудита
        отклоняется до старта (аудит устойчивости, находки 1.6a, 1.6b, 1.6d).
        """
        configs_dir = Path(base_dir) / "doc_configs" if base_dir else _DOC_CONFIGS_DIR
        result = []
        if not configs_dir.exists():
            return result
        for config_dir in sorted(configs_dir.iterdir()):
            config_file = config_dir / "config.json"
            if not config_file.exists():
                continue
            entry: Dict[str, Any] = {"doc_type": config_dir.name, "doc_title": "(ошибка чтения config.json)"}
            reason = _validate_doc_type_dir(config_dir, entry)
            if reason:
                entry["broken"] = True
                entry["broken_reason"] = reason
            result.append(entry)
        return result
# END_AUDIT_ENGINE
