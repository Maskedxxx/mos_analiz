# START_MODULE_CONTRACT
# PURPOSE: Generic-движок аудита для 20 docx/pptx doc_types через multi_rule LLM-путь. Используется когда `AuditConfig.parser_by_ext` задан (НЕ для special-движков kpsc/kartochka/drivers/plan_grafik — у тех свои runner-ы в `src/doc_type_validators/`).
# INPUTS: doc_type (имя папки в doc_configs/), target file. Конфиг через `load_audit_config`. LLM через `run_multi_rule_audit`.
# OUTPUTS: `AuditResult` с violations + Excel-отчёт + JSON-логи в session_dir.
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
# END_IMPORTS


# START_PATHS
# PURPOSE: Локальные пути модуля. parents[2] = repo root (engine.py лежит в src/audit/).
_PROJECT_ROOT = Path(__file__).resolve().parents[2]
_DOC_CONFIGS_DIR = _PROJECT_ROOT / "doc_configs"
# END_PATHS


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
            print_prompts: режим отладки — выводить промпты без LLM.
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
        _emit("audit_start", {"doc_type": self.doc_type, "filename": Path(target_path).name})
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
        # Активный LLM-путь — multi_rule: парсер отдаёт raw_text, промпт собирается
        # по sections.json + rules_multi.json / rules_methodology.json.
        doc_configs_dir = self.config.config_dir.parent
        mr_config = load_multi_rule_config(doc_configs_dir, self.doc_type)
        if mr_config is None:
            raise ValueError(
                f"Для doc_type={self.doc_type} не найдены sections.json + rules_multi.json. "
                f"Без них multi-rule аудит запустить нельзя; legacy single-rule путь удалён."
            )
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
        )
        violations = mr_result["violations"]
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
            )
            violations.extend(meth_result["violations"])
            self.logger.log(f"   Methodology usage: prompt={meth_result['usage']['prompt_tokens']} completion={meth_result['usage']['completion_tokens']} total={meth_result['usage']['total_tokens']}")
            self.logger.log(f"   Methodology violations: {len(meth_result['violations'])}")
        _all_multi_rules = [{**r, "layer": "base"} for r in mr_config["rules"]]
        if meth_config is not None:
            _all_multi_rules.extend(({**r, "layer": "methodology"} for r in meth_config["rules"]))
        _emit("checking_rules_done", {"violations": len(violations)})
        self.logger.log_final_results(violations)
        print(json.dumps(violations, ensure_ascii=False, indent=2))
        if out_xlsx:
            xlsx_path = out_xlsx
        else:
            xlsx_path = str(session_path / "audit_result.xlsx")
        save_to_excel(violations, xlsx_path, multi_rules=_all_multi_rules)
        self.logger.log(f"📊 Excel сохранён: {xlsx_path}")
        duration = time.time() - start_time
        self.logger.log(f"{'=' * 60}")
        self.logger.log(f"📊 Итого нарушений: {len(violations)}")
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
            duration_sec=duration, rules_checked=len(_all_multi_rules), target_path=target_path,
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
    def list_doc_types(base_dir: Optional[str] = None) -> List[Dict[str, str]]:
        """
        Список доступных типов документов.

        Сканирует doc_configs/ и возвращает список {doc_type, doc_title}.
        """
        configs_dir = Path(base_dir) / "doc_configs" if base_dir else _DOC_CONFIGS_DIR
        result = []
        if not configs_dir.exists():
            return result
        for config_dir in sorted(configs_dir.iterdir()):
            config_file = config_dir / "config.json"
            if not config_file.exists():
                continue
            try:
                with open(config_file, "r", encoding="utf-8") as f:
                    data = json.load(f)
                result.append({"doc_type": data.get("doc_type", config_dir.name), "doc_title": data.get("doc_title", "")})
            except (json.JSONDecodeError, KeyError):
                result.append({"doc_type": config_dir.name, "doc_title": "(ошибка чтения config.json)"})
        return result
# END_AUDIT_ENGINE
