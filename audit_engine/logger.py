#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Логгер пайплайна аудита.

Сохраняет все артефакты сессии:
- prompts/       — промпты для каждого правила
- responses/     — ответы LLM и non-LLM результаты
- parsed_docs/   — распарсенные документы в JSON
- pipeline.log   — главный лог выполнения
- final_results.json — итоговые результаты
"""

import json
import sys
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List


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

        # Создаём поддиректории
        self.prompts_dir = log_dir / "prompts"
        self.responses_dir = log_dir / "responses"
        self.parsed_dir = log_dir / "parsed_docs"

        self.prompts_dir.mkdir(exist_ok=True)
        self.responses_dir.mkdir(exist_ok=True)
        self.parsed_dir.mkdir(exist_ok=True)

        # Основной лог-файл
        self.main_log = log_dir / "pipeline.log"
        self._init_main_log(doc_type)

    def _init_main_log(self, doc_type: str):
        """Инициализирует главный лог-файл."""
        with open(self.main_log, 'w', encoding='utf-8') as f:
            f.write(f"{'='*80}\n")
            f.write(f"AUDIT PIPELINE LOG — {doc_type or 'Universal Audit Engine'}\n")
            f.write(f"Started: {datetime.now().isoformat()}\n")
            f.write(f"{'='*80}\n\n")

    def log(self, message: str):
        """Записывает сообщение в главный лог и stderr."""
        timestamp = datetime.now().strftime("%H:%M:%S")
        with open(self.main_log, 'a', encoding='utf-8') as f:
            f.write(f"[{timestamp}] {message}\n")
        print(f"[{timestamp}] {message}", file=sys.stderr)

    def log_parsed_doc(self, doc: Dict[str, Any], name: str):
        """Сохраняет распарсенный документ в JSON."""
        filepath = self.parsed_dir / f"{name}.json"
        with open(filepath, 'w', encoding='utf-8') as f:
            json.dump(doc, f, ensure_ascii=False, indent=2)
        self.log(f"📄 Распарсенный документ сохранён: {filepath}")

    def log_rule_prompt(self, rule_index: int, system_prompt: str, user_prompt: str):
        """Сохраняет промпт для правила."""
        filepath = self.prompts_dir / f"rule_{rule_index:02d}_prompt.txt"
        with open(filepath, 'w', encoding='utf-8') as f:
            f.write(f"{'='*80}\n")
            f.write(f"RULE #{rule_index} - PROMPT\n")
            f.write(f"{'='*80}\n\n")
            f.write(f"--- SYSTEM PROMPT ---\n")
            f.write(system_prompt)
            f.write(f"\n\n--- USER PROMPT ---\n")
            f.write(user_prompt)
        self.log(f"📝 Промпт для правила #{rule_index} сохранён: {filepath}")

    def log_rule_response(self, rule_index: int, raw_response: str, parsed_violations: List[Dict]):
        """Сохраняет ответ LLM для правила."""
        filepath = self.responses_dir / f"rule_{rule_index:02d}_response.txt"
        with open(filepath, 'w', encoding='utf-8') as f:
            f.write(f"{'='*80}\n")
            f.write(f"RULE #{rule_index} - LLM RESPONSE\n")
            f.write(f"{'='*80}\n\n")
            f.write(f"--- RAW RESPONSE ---\n")
            f.write(raw_response)
            f.write(f"\n\n--- PARSED VIOLATIONS ---\n")
            f.write(json.dumps(parsed_violations, ensure_ascii=False, indent=2))
        self.log(f"✅ Ответ для правила #{rule_index} сохранён: {filepath}")

    def log_non_llm_result(self, rule_index: int, violations: List[Dict]):
        """Сохраняет результат non-LLM проверки."""
        filepath = self.responses_dir / f"rule_{rule_index:02d}_non_llm.txt"
        with open(filepath, 'w', encoding='utf-8') as f:
            f.write(f"{'='*80}\n")
            f.write(f"RULE #{rule_index} - NON-LLM CHECK\n")
            f.write(f"{'='*80}\n\n")
            f.write(f"--- RESULT ---\n")
            f.write(json.dumps(violations, ensure_ascii=False, indent=2))
        self.log(f"✅ Результат non-LLM правила #{rule_index} сохранён: {filepath}")

    def log_final_results(self, violations: List[Dict]):
        """Сохраняет финальные результаты."""
        filepath = self.log_dir / "final_results.json"
        with open(filepath, 'w', encoding='utf-8') as f:
            json.dump(violations, f, ensure_ascii=False, indent=2)
        self.log(f"📊 Финальные результаты сохранены: {filepath}")

    def log_error(self, rule_index: int, error: str):
        """Логирует ошибку для правила."""
        filepath = self.responses_dir / f"rule_{rule_index:02d}_error.txt"
        with open(filepath, 'w', encoding='utf-8') as f:
            f.write(f"ERROR for rule #{rule_index}:\n{error}")
        self.log(f"❌ ОШИБКА для правила #{rule_index}: {error}")
