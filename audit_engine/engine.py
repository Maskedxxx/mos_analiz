#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
AuditEngine — оркестратор аудита документов.

Полный цикл:
1. Загрузка конфигов из doc_configs/<doc_type>/
2. Vision-парсинг целевого документа и шаблона
3. Кэширование шаблона (SHA256)
4. Применение препроцессоров
5. Параллельные non-LLM + LLM проверки
6. Сохранение результатов (Excel + JSON + логи)
"""

import hashlib
import json
import sys
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional

from .models import AuditConfig, AuditResult, RuleSpec, load_audit_config, load_rules
from .logger import PipelineLogger
from .context_builder import build_user_prompt
from .llm_client import call_llm, parse_json_response
from .excel_reporter import save_to_excel
from .non_llm_checks.registry import get_check
from .preprocessors.registry import get_preprocessors


class AuditEngine:
    """
    Единый движок аудита документов.

    Использование:
        engine = AuditEngine("presentation_eu")
        result = engine.run("document.pptx")

    Или с кастомным base_dir:
        engine = AuditEngine("prikaz_ic", base_dir="/path/to/project")
    """

    def __init__(
        self,
        doc_type: str,
        base_dir: Optional[str] = None,
        config_dir: Optional[str] = None
    ):
        """
        Инициализация движка.

        Args:
            doc_type: тип документа (имя папки в doc_configs/)
            base_dir: корневая директория проекта (по умолчанию — родитель audit_engine/)
            config_dir: путь к папке конфигов (переопределяет base_dir + doc_configs)
        """
        self.doc_type = doc_type

        # Определяем base_dir
        if base_dir:
            self.base_dir = Path(base_dir)
        else:
            # Родитель audit_engine/ — корень проекта
            self.base_dir = Path(__file__).parent.parent

        # Загружаем конфиг
        if config_dir:
            self.config_path = Path(config_dir)
        else:
            self.config_path = self.base_dir / "doc_configs" / doc_type

        self.config: AuditConfig = load_audit_config(self.config_path)

        # Загружаем системный промпт
        self.system_prompt = self._load_system_prompt()

        # Загружаем препроцессоры для этого типа документа
        self.preprocessors = get_preprocessors(doc_type)

        # Логгер (создаётся при run())
        self.logger: Optional[PipelineLogger] = None

    def _load_system_prompt(self) -> str:
        """Загружает системный промпт из файла."""
        prompt_file = (
            Path(__file__).parent / "system_prompts" / self.config.system_prompt
        )
        if prompt_file.exists():
            return prompt_file.read_text(encoding='utf-8')

        # Fallback на default.txt
        default_file = Path(__file__).parent / "system_prompts" / "default.txt"
        if default_file.exists():
            return default_file.read_text(encoding='utf-8')

        return "Ты — строгий аудитор документов."

    def run(
        self,
        target_path: str,
        template_path: Optional[str] = None,
        model: Optional[str] = None,
        temperature: Optional[float] = None,
        rule_filter: Optional[int] = None,
        parse_only: bool = False,
        no_cache: bool = False,
        print_prompts: bool = False,
        session_dir: Optional[str] = None,
        chunk_filter: Optional[str] = None,
        out_xlsx: Optional[str] = None,
        secondary_path: Optional[str] = None,
        progress_callback: Optional[callable] = None
    ) -> AuditResult:
        """
        Запуск полного цикла аудита.

        Args:
            target_path: путь к целевому документу
            template_path: путь к шаблону (если не указан — из конфига)
            model: модель LLM (переопределяет конфиг)
            temperature: температура (переопределяет конфиг)
            rule_filter: проверить только одно правило
            parse_only: режим только парсинга
            no_cache: не использовать кэш шаблона
            print_prompts: режим отладки — выводить промпты без LLM
            session_dir: директория для логов
            chunk_filter: парсить только указанный чанк
            out_xlsx: путь для сохранения Excel
            secondary_path: путь к вторичному файлу (XLSX для multi-file аудитов)
            progress_callback: колбэк прогресса для веб-UI (type: str, data: dict)

        Returns:
            AuditResult с нарушениями и статистикой
        """
        start_time = time.time()

        # Хелпер для отправки прогресса (если есть колбэк)
        def _emit(event_type: str, data: dict = None):
            if progress_callback:
                progress_callback(event_type, data or {})

        # Параметры из конфига с возможностью переопределения
        model = model or self.config.model
        temperature = temperature if temperature is not None else self.config.temperature

        # Инициализируем сессию
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
        self.logger.log(f"   Парсер: {self.config.parser}")
        if self.config.parser == "paddle":
            self.logger.log(f"   Layout: {self.config.paddle_layout_model or 'Heron-101 (дефолт)'}")
            self.logger.log(f"   VLM: {self.config.paddle_vlm_model or 'PaddleOCR-VL-1.5 (авто)'}")
            self.logger.log(f"   VLM URL: {self.config.ocr_base_url}")
        if self.config.parser == "ocr":
            self.logger.log(f"   OCR URL: {self.config.ocr_base_url}")
        if self.config.llm_base_url:
            self.logger.log(f"   LLM URL: {self.config.llm_base_url}")

        # Проверяем наличие целевого файла
        if not Path(target_path).exists():
            self.logger.log(f"❌ Целевой документ не найден: {target_path}")
            raise FileNotFoundError(f"Целевой документ не найден: {target_path}")

        # Предупреждение: конфиг требует secondary, но не передан
        if self.config.secondary_file and not secondary_path:
            self.logger.log(f"⚠️ Конфиг требует вторичный файл ({self.config.secondary_file.type}), но --secondary не указан")

        # Проверяем наличие вторичного файла
        if secondary_path and not Path(secondary_path).exists():
            self.logger.log(f"❌ Вторичный файл не найден: {secondary_path}")
            raise FileNotFoundError(f"Вторичный файл не найден: {secondary_path}")

        # Загружаем правила
        self.logger.log(f"📋 Загрузка правил...")
        all_rules = load_rules(str(self.config.rules_path))
        self.logger.log(f"   Загружено {len(all_rules)} правил")
        _emit("audit_start", {"doc_type": self.doc_type, "filename": Path(target_path).name, "total_rules": len(all_rules)})

        # Фильтр по правилу
        rules = all_rules
        chunks_to_parse = None

        if rule_filter is not None:
            rules = [r for r in all_rules if r.index == rule_filter]
            if not rules:
                self.logger.log(f"❌ Правило #{rule_filter} не найдено")
                raise ValueError(f"Правило #{rule_filter} не найдено в {self.doc_type}")
            rule = rules[0]
            chunks_to_parse = [rule.scope] if isinstance(rule.scope, str) else list(rule.scope)
            self.logger.log(f"   ⚠️ Фильтр: только правило #{rule_filter}")

        if chunk_filter:
            chunks_to_parse = [chunk_filter]

        # Определяем тип загруженного файла для doc_types с secondary_file
        # Если загружен xlsx вместо docx — парсим как вторичный файл
        target_ext = Path(target_path).suffix.lower()
        _is_secondary_upload = False

        if (self.config.secondary_file
                and target_ext == f".{self.config.secondary_file.type}"
                and not secondary_path):
            # Пользователь загрузил вторичный файл (xlsx) как основной
            _is_secondary_upload = True
            self.logger.log(f"📊 Загружен {target_ext} — парсим как вторичный файл")
            _emit("parsing_target", {})
            target_doc = self._parse_secondary(target_path)
            _emit("parsing_target_done", {})
        else:
            # Стандартный парсинг основного документа
            _emit("parsing_target", {})
            target_doc = self._parse_document(target_path, chunks_to_parse, session_path / "vision_target")
            _emit("parsing_target_done", {})

            # Парсинг вторичного файла и merge в target_doc
            if secondary_path and self.config.secondary_file:
                secondary_chunks = self._parse_secondary(secondary_path)
                target_doc.update(secondary_chunks)
                self.logger.log(f"   📊 Merged {len(secondary_chunks)} чанков из вторичного файла")

        # Получаем шаблон
        _emit("parsing_template", {})
        template_doc = self._get_template(
            template_path=template_path,
            no_cache=no_cache,
            chunks_to_parse=chunks_to_parse,
            session_dir=session_path
        )
        _emit("parsing_template_done", {})

        # Режим --parse-only
        if parse_only:
            self.logger.log(f"✅ Режим --parse-only: парсинг завершён")
            print(f"\n{'='*60}")
            print("TARGET:")
            print(json.dumps(target_doc, ensure_ascii=False, indent=2))
            print(f"\n{'='*60}")
            print("TEMPLATE:")
            print(json.dumps(template_doc, ensure_ascii=False, indent=2))
            return AuditResult(
                doc_type=self.doc_type,
                session_dir=session_path,
                target_path=target_path,
                duration_sec=time.time() - start_time
            )

        # Фильтруем правила: пропускаем те, чьи scope-чанки отсутствуют в target_doc
        # (актуально для doc_types с secondary_file — если загружен только один файл)
        available_keys = set(target_doc.keys())
        filtered_rules = []
        for r in rules:
            scopes = [r.scope] if isinstance(r.scope, str) else list(r.scope)
            # Правило с scope "имя_файла" всегда доступно
            if all(s == "имя_файла" or s in available_keys for s in scopes):
                filtered_rules.append(r)
            else:
                missing = [s for s in scopes if s != "имя_файла" and s not in available_keys]
                self.logger.log(f"   ⏭️ Пропуск правила #{r.index} ({r.title}): нет чанков {missing}")

        if len(filtered_rules) < len(rules):
            self.logger.log(f"   📋 Доступно {len(filtered_rules)} из {len(rules)} правил (остальные пропущены)")
        rules = filtered_rules

        # Сохраняем сводку правил
        rules_summary = [
            {"index": r.index, "title": r.title, "llm": r.llm, "compare": r.compare}
            for r in rules
        ]
        with open(session_path / "rules_summary.json", 'w', encoding='utf-8') as f:
            json.dump(rules_summary, f, ensure_ascii=False, indent=2)

        # Выполняем проверки
        self.logger.log(f"🔍 Запуск проверок...")
        violations = self._run_checks(
            rules=rules,
            target_doc=target_doc,
            template_doc=template_doc,
            model=model,
            temperature=temperature,
            print_prompts=print_prompts,
            progress_callback=progress_callback
        )

        # Сохраняем результаты
        self.logger.log_final_results(violations)
        print(json.dumps(violations, ensure_ascii=False, indent=2))

        # Excel
        if out_xlsx:
            xlsx_path = out_xlsx
        else:
            xlsx_path = str(session_path / "audit_result.xlsx")
        save_to_excel(violations, xlsx_path)
        self.logger.log(f"📊 Excel сохранён: {xlsx_path}")

        # Итоговая статистика
        duration = time.time() - start_time
        self.logger.log(f"{'='*60}")
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
            violations=violations,
            doc_type=self.doc_type,
            session_dir=session_path,
            duration_sec=duration,
            rules_checked=len(rules),
            target_path=target_path
        )

    def _parse_document(
        self,
        file_path: str,
        chunks_to_parse: Optional[List[str]],
        vision_log_dir: Path
    ) -> Dict[str, Any]:
        """
        Парсит документ через выбранный парсер (Vision или OCR).

        Выбор парсера определяется config.parser:
        - "vision" → VisionParser (облачный GPT Vision)
        - "ocr" → OcrParser (локальный HunyuanOCR + сборка по страницам)
        """
        chunk_filter_arg = (
            chunks_to_parse[0]
            if chunks_to_parse and len(chunks_to_parse) == 1
            else None
        )

        # Автовыбор парсера для DOCX: если parser=paddle, но файл .docx → docx_parser
        file_ext = Path(file_path).suffix.lower()
        effective_parser = self.config.parser
        if effective_parser == "paddle" and file_ext == ".docx":
            effective_parser = "docx"
            self.logger.log(f"📄 Файл .docx — переключение на docx-парсер (без OCR)")

        if effective_parser == "docx":
            # Прямое извлечение текста из DOCX через python-docx (без OCR)
            self.logger.log(f"📄 DOCX-парсинг: {file_path}...")
            from .docx_parser import parse_docx
            doc = parse_docx(
                file_path,
                str(self.config.chunks_vision_path),
                chunk_filter=chunk_filter_arg,
            )
        elif effective_parser == "pptx":
            # Прямое извлечение текста из PPTX через python-pptx (без OCR)
            self.logger.log(f"📄 PPTX-парсинг: {file_path}...")
            from .pptx_parser import parse_pptx
            doc = parse_pptx(
                file_path,
                str(self.config.chunks_vision_path),
                chunk_filter=chunk_filter_arg,
            )
        elif effective_parser == "paddle":
            # Layout-aware OCR: Heron-101 + PaddleOCR-VL-1.5
            self.logger.log(f"📄 Paddle-парсинг: {file_path}...")
            try:
                from .paddle_parser import PaddleParser
            except ImportError as e:
                self.logger.log(f"❌ Не удалось импортировать PaddleParser: {e}")
                raise ImportError(f"PaddleParser недоступен: {e}") from e

            parser = PaddleParser(config=self.config, log_dir=str(vision_log_dir))
            doc = parser.parse(file_path, chunk_filter=chunk_filter_arg)
        elif self.config.parser == "ocr":
            # Локальный OCR-парсер
            self.logger.log(f"📄 OCR-парсинг: {file_path}...")
            try:
                from .ocr_parser import OcrParser
            except ImportError as e:
                self.logger.log(f"❌ Не удалось импортировать OcrParser: {e}")
                raise ImportError(f"OcrParser недоступен: {e}") from e

            parser = OcrParser(config=self.config, log_dir=str(vision_log_dir))
            doc = parser.parse(file_path, chunk_filter=chunk_filter_arg)
        else:
            # Облачный Vision-парсер (по умолчанию)
            self.logger.log(f"📄 Vision-парсинг: {file_path}...")
            try:
                from .vision_parser import VisionParser
            except ImportError as e:
                self.logger.log(f"❌ Не удалось импортировать VisionParser: {e}")
                self.logger.log(f"   Убедитесь, что установлены зависимости: pip install pdf2image tenacity Pillow")
                raise ImportError(f"VisionParser недоступен: {e}") from e

            vision_parser = VisionParser(
                str(self.config.chunks_vision_path),
                log_dir=str(vision_log_dir)
            )
            doc = vision_parser.parse(file_path, chunk_filter=chunk_filter_arg)

        self.logger.log_parsed_doc(doc, Path(file_path).stem)
        return doc

    def _get_template(
        self,
        template_path: Optional[str],
        no_cache: bool,
        chunks_to_parse: Optional[List[str]],
        session_dir: Path
    ) -> Dict[str, Any]:
        """
        Получает шаблон: из кэша или через Vision-парсинг.

        Кэширование по SHA256:
        - Если template_cached.json существует и хэш совпадает — используем кэш
        - Иначе парсим через Vision и сохраняем кэш
        """
        # Определяем путь к шаблону
        if template_path:
            tpl_path = Path(template_path)
        elif self.config.template_path:
            tpl_path = self.config.template_path
        else:
            self.logger.log("⚠️ Шаблон не указан, пропускаем")
            return {}

        if not tpl_path.exists():
            self.logger.log(f"❌ Шаблон не найден: {tpl_path}")
            return {}

        self.logger.log(f"📄 Шаблон: {tpl_path}")

        # Вычисляем хэш файла шаблона
        file_hash = self._compute_file_hash(tpl_path)

        # Проверяем кэш
        cache_path = tpl_path.parent / "template_cached.json"

        if not no_cache and cache_path.exists():
            try:
                with open(cache_path, 'r', encoding='utf-8') as f:
                    cached = json.load(f)

                # Проверяем совпадение хэша файла И типа парсера
                cache_parser = cached.get("_parser", "vision")
                if cached.get("_hash") == file_hash and cache_parser == self.config.parser:
                    self.logger.log(f"✅ Используем кэш шаблона: {cache_path}")
                    # Убираем служебные поля
                    cached.pop("_hash", None)
                    cached.pop("_cached_at", None)
                    cached.pop("_parser", None)
                    return cached
                elif cached.get("_hash") != file_hash:
                    self.logger.log(f"⚠️ Кэш устарел (хэш изменился), перепарсинг...")
                else:
                    self.logger.log(f"⚠️ Кэш от другого парсера ({cache_parser}→{self.config.parser}), перепарсинг...")
            except (json.JSONDecodeError, KeyError):
                self.logger.log(f"⚠️ Кэш повреждён, перепарсинг...")

        # Парсим шаблон через Vision
        self.logger.log(f"🔮 Vision-парсинг шаблона...")
        template_doc = self._parse_document(
            str(tpl_path),
            chunks_to_parse,
            session_dir / "vision_template"
        )

        # Сохраняем кэш (с типом парсера для инвалидации при смене vision↔ocr)
        cache_data = dict(template_doc)
        cache_data["_hash"] = file_hash
        cache_data["_parser"] = self.config.parser
        cache_data["_cached_at"] = datetime.now().isoformat()

        with open(cache_path, 'w', encoding='utf-8') as f:
            json.dump(cache_data, f, ensure_ascii=False, indent=2)
        self.logger.log(f"💾 Кэш шаблона сохранён: {cache_path}")

        return template_doc

    @staticmethod
    def _compute_file_hash(file_path: Path) -> str:
        """Вычисляет SHA256 хэш файла."""
        sha256 = hashlib.sha256()
        with open(file_path, 'rb') as f:
            for chunk in iter(lambda: f.read(8192), b''):
                sha256.update(chunk)
        return sha256.hexdigest()

    def _parse_secondary(self, file_path: str) -> Dict[str, Any]:
        """
        Парсит вторичный файл через зарегистрированный парсер.

        Добавляет префикс из config.secondary_file.chunk_prefix к ключам,
        а также ключ <prefix>имя_файла с именем файла.
        """
        from .parsers import get_parser

        spec = self.config.secondary_file
        parser_fn = get_parser(spec.parser)
        raw_chunks = parser_fn(file_path)

        # Добавляем префикс ко всем ключам
        prefixed = {f"{spec.chunk_prefix}{k}": v for k, v in raw_chunks.items()}

        # Добавляем имя файла
        prefixed[f"{spec.chunk_prefix}имя_файла"] = Path(file_path).name

        self.logger.log_parsed_doc(prefixed, f"secondary_{spec.type}")
        return prefixed

    def _run_checks(
        self,
        rules: List[RuleSpec],
        target_doc: Dict[str, Any],
        template_doc: Dict[str, Any],
        model: str,
        temperature: float,
        print_prompts: bool = False,
        progress_callback: Optional[callable] = None
    ) -> List[Dict[str, Any]]:
        """
        Выполняет все проверки (non-LLM + LLM).

        Non-LLM правила выполняются синхронно.
        LLM правила выполняются параллельно через ThreadPoolExecutor.
        """
        all_violations: List[Dict[str, Any]] = []
        rules_done = 0
        total_rules = len(rules)

        # === Non-LLM правила ===
        for spec in rules:
            if not spec.llm:
                self.logger.log(f"🔧 Запуск non-LLM правила #{spec.index}: {spec.title}")
                check_fn = get_check(self.doc_type, spec.index)
                if check_fn:
                    try:
                        violations = check_fn(target_doc, self.config)
                    except Exception as e:
                        self.logger.log(f"   ❌ Ошибка в non-LLM проверке #{spec.index}: {e}")
                        violations = []
                else:
                    self.logger.log(f"   ⚠️ Нет зарегистрированной проверки для {self.doc_type}#{spec.index}")
                    violations = []

                all_violations.extend(violations)
                self.logger.log_non_llm_result(spec.index, violations)
                rules_done += 1
                if progress_callback:
                    progress_callback("rule_done", {
                        "rule_index": spec.index, "rule_title": spec.title,
                        "current": rules_done, "total": total_rules,
                        "violations_count": len(violations)
                    })

                if print_prompts:
                    print(f"\n{'='*60}")
                    print(f"[NON-LLM] Правило #{spec.index}: {spec.title}")
                    print(f"Результат: {violations}")

        # === LLM правила ===
        llm_rules = [r for r in rules if r.llm]

        if print_prompts:
            # Режим отладки — только промпты, без вызова LLM
            for spec in llm_rules:
                user_prompt = build_user_prompt(spec, target_doc, template_doc, self.preprocessors)
                print(f"\n{'='*60}")
                print(f"[LLM] Правило #{spec.index}: {spec.title}")
                print(f"{'='*60}")
                print("\n--- SYSTEM PROMPT ---")
                print(self.system_prompt)
                print("\n--- USER PROMPT ---")
                print(user_prompt)

                self.logger.log_rule_prompt(spec.index, self.system_prompt, user_prompt)
            return all_violations

        # Параллельное выполнение LLM-проверок
        max_workers = self.config.max_workers
        self.logger.log(f"🚀 Запуск {len(llm_rules)} LLM-проверок (max_workers={max_workers})")

        with ThreadPoolExecutor(max_workers=max_workers) as executor:
            futures = {
                executor.submit(
                    self._run_single_llm_check,
                    spec,
                    target_doc,
                    template_doc,
                    model,
                    temperature
                ): spec
                for spec in llm_rules
            }

            for future in as_completed(futures):
                spec = futures[future]
                try:
                    violations = future.result()
                    all_violations.extend(violations)
                    self.logger.log(f"✅ Правило #{spec.index} проверено, нарушений: {len(violations)}")
                    rules_done += 1
                    if progress_callback:
                        progress_callback("rule_done", {
                            "rule_index": spec.index, "rule_title": spec.title,
                            "current": rules_done, "total": total_rules,
                            "violations_count": len(violations)
                        })
                except Exception as e:
                    error_msg = str(e)
                    self.logger.log_error(spec.index, error_msg)
                    print(f"[ERROR] Правило #{spec.index}: {e}", file=sys.stderr)

        return sorted(all_violations, key=lambda x: x.get("rule_index", 0))

    def _run_single_llm_check(
        self,
        spec: RuleSpec,
        target_doc: Dict[str, Any],
        template_doc: Dict[str, Any],
        model: str,
        temperature: float
    ) -> List[Dict[str, Any]]:
        """Проверка одного правила через LLM."""
        user_prompt = build_user_prompt(spec, target_doc, template_doc, self.preprocessors)

        messages = [
            {"role": "system", "content": self.system_prompt},
            {"role": "user", "content": user_prompt}
        ]

        self.logger.log_rule_prompt(spec.index, self.system_prompt, user_prompt)

        raw_response = call_llm(
            messages, model, temperature,
            base_url=self.config.llm_base_url,
            max_tokens=self.config.llm_max_tokens,
            reasoning_effort=self.config.reasoning_effort,
            seed=self.config.llm_seed
        )

        violations = parse_json_response(raw_response, spec.index, spec.title)

        # Гарантируем наличие rule_index/rule_title в каждом нарушении
        for obj in violations:
            obj.setdefault("rule_index", spec.index)
            obj.setdefault("rule_title", spec.title)

        self.logger.log_rule_response(spec.index, raw_response, violations)

        return violations

    @staticmethod
    def list_doc_types(base_dir: Optional[str] = None) -> List[Dict[str, str]]:
        """
        Список доступных типов документов.

        Сканирует doc_configs/ и возвращает список {doc_type, doc_title}.
        """
        if base_dir:
            configs_dir = Path(base_dir) / "doc_configs"
        else:
            configs_dir = Path(__file__).parent.parent / "doc_configs"

        result = []
        if not configs_dir.exists():
            return result

        for config_dir in sorted(configs_dir.iterdir()):
            config_file = config_dir / "config.json"
            if config_file.exists():
                try:
                    with open(config_file, 'r', encoding='utf-8') as f:
                        data = json.load(f)
                    result.append({
                        "doc_type": data.get("doc_type", config_dir.name),
                        "doc_title": data.get("doc_title", ""),
                    })
                except (json.JSONDecodeError, KeyError):
                    result.append({
                        "doc_type": config_dir.name,
                        "doc_title": "(ошибка чтения config.json)",
                    })

        return result
