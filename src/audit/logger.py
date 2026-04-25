# START_MODULE_CONTRACT
# PURPOSE: Логгер сессии аудита. Пишет pipeline.log + сохраняет распарсенные документы и финальные результаты в JSON-файлы внутри session_dir.
# INPUTS: session_dir (Path), doc_type (str). Сообщения от runtime через `log()`, `log_parsed_doc()`, `log_final_results()`.
# OUTPUTS: Файлы в session_dir: `pipeline.log`, `parsed_docs/<name>.json`, `final_results.json`.
# KEYWORDS: logger, pipeline, session, audit.
# LINKS: src/audit/engine.py (использует через AuditEngine.logger), src/doc_type_validators/*.py (special-runner-ы могут переиспользовать).
# RATIONALE: Тонкий вспомогательный модуль для логирования по сессии. Не зависит от движка/моделей — берёт только Path и базовые типы.
# END_MODULE_CONTRACT

from __future__ import annotations

# START_IMPORTS
import json
import sys
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List
# END_IMPORTS


# START_PIPELINE_LOGGER
class PipelineLogger:
    """
    Логгер для всего пайплайна аудита.
    Сохраняет все этапы в отдельные файлы для отладки.
    """

    def __init__(self, log_dir: Path, doc_type: str = ""):
        """
        Инициализация логгера.

        Args:
            log_dir: директория для логов сессии
            doc_type: тип документа (для заголовка лога)
        """
        self.log_dir = log_dir
        self.log_dir.mkdir(parents=True, exist_ok=True)
        self.parsed_dir = log_dir / "parsed_docs"
        self.parsed_dir.mkdir(exist_ok=True)
        self.main_log = log_dir / "pipeline.log"
        self._init_main_log(doc_type)

    def _init_main_log(self, doc_type: str) -> None:
        """Инициализирует главный лог-файл."""
        with open(self.main_log, "w", encoding="utf-8") as f:
            f.write(f"{'=' * 80}\n")
            f.write(f"AUDIT PIPELINE LOG — {doc_type or 'Universal Audit Engine'}\n")
            f.write(f"Started: {datetime.now().isoformat()}\n")
            f.write(f"{'=' * 80}\n\n")

    def log(self, message: str) -> None:
        """Записывает сообщение в главный лог и stderr."""
        timestamp = datetime.now().strftime("%H:%M:%S")
        with open(self.main_log, "a", encoding="utf-8") as f:
            f.write(f"[{timestamp}] {message}\n")
        print(f"[{timestamp}] {message}", file=sys.stderr)

    def log_parsed_doc(self, doc: Dict[str, Any], name: str) -> None:
        """Сохраняет распарсенный документ в JSON."""
        filepath = self.parsed_dir / f"{name}.json"
        with open(filepath, "w", encoding="utf-8") as f:
            json.dump(doc, f, ensure_ascii=False, indent=2)
        self.log(f"📄 Распарсенный документ сохранён: {filepath}")

    def log_final_results(self, violations: List[Dict]) -> None:
        """Сохраняет финальные результаты."""
        filepath = self.log_dir / "final_results.json"
        with open(filepath, "w", encoding="utf-8") as f:
            json.dump(violations, f, ensure_ascii=False, indent=2)
        self.log(f"📊 Финальные результаты сохранены: {filepath}")
# END_PIPELINE_LOGGER
