#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
LLM-аудит документов "Приказ о проведении выхода" + "График обхода ОМ".

Проверяет DOCX и XLSX файлы по правилам из tz_prikaz_vyhod.json.
Поддерживает три типа сравнения:
- template: сверка DOCX с шаблоном
- target_only: проверка только целевого файла
- cross_file: сверка между DOCX и XLSX

Все артефакты сессии сохраняются в:
logs_result/session_YYYYMMDD_HHMMSS/
├── prompts/           # Промпты для каждого правила
├── responses/         # Ответы LLM
├── parsed_docs/       # Распарсенные документы
├── pipeline.log       # Лог выполнения
├── final_results.json # JSON с нарушениями
├── rules_summary.json # Сводка по правилам
└── audit_result.xlsx  # Excel-отчёт

Использование:
    python run_prikaz_vyhod_audit.py --docx <путь> --xlsx <путь> --template <путь>
"""

import argparse
import json
import os
import re
import sys
from concurrent.futures import ThreadPoolExecutor, as_completed
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional

import pandas as pd
from openai import OpenAI

# Добавляем путь к модулям
sys.path.insert(0, str(Path(__file__).parent))
from parser_prikaz_vyhod_docx import parse_prikaz_vyhod
from parser_grafik_xlsx import parse_grafik_obhod


# === КОНФИГУРАЦИЯ ===
PROJECT_DIR = Path(__file__).parent.parent
DEFAULT_RULES = PROJECT_DIR / "config" / "tz_prikaz_vyhod.json"
DEFAULT_MODEL = "gpt-4.1-mini"
MAX_WORKERS = 4


# ============================================================================
# ПРЕПРОЦЕССИНГ ДЛЯ ПРАВИЛА #3
# ============================================================================

def normalize_text_for_rule3(text: str) -> str:
    """
    Нормализует текст для правила #3, заменяя плейсхолдеры на унифицированные метки.
    """
    # Нормализуем п.1: должность + ФИО организатора
    text = re.sub(
        r'(1\.\s*Назначить организатором проведения обхода на производственной площадке\s*).+?\.',
        r'\1[ДОЛЖНОСТЬ_ФИО].',
        text,
        flags=re.DOTALL
    )

    # Нормализуем п.2: должность + ФИО секретаря
    text = re.sub(
        r'(2\.\s*Назначить секретарем проведения обхода на производственной площадке\s*).+?\.',
        r'\1[ДОЛЖНОСТЬ_ФИО].',
        text,
        flags=re.DOTALL
    )

    return text


# ============================================================================
# ЛОГИРОВАНИЕ
# ============================================================================

class PipelineLogger:
    """Логгер для пайплайна аудита. Сохраняет все этапы для отладки."""

    def __init__(self, log_dir: Path):
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
        self._init_main_log()

    def _init_main_log(self):
        with open(self.main_log, 'w', encoding='utf-8') as f:
            f.write(f"{'='*80}\n")
            f.write(f"AUDIT PIPELINE LOG\n")
            f.write(f"Started: {datetime.now().isoformat()}\n")
            f.write(f"{'='*80}\n\n")

    def log(self, message: str):
        timestamp = datetime.now().strftime("%H:%M:%S")
        with open(self.main_log, 'a', encoding='utf-8') as f:
            f.write(f"[{timestamp}] {message}\n")

    def log_parsed_doc(self, doc: Dict[str, Any], name: str):
        filepath = self.parsed_dir / f"{name}.json"
        with open(filepath, 'w', encoding='utf-8') as f:
            json.dump(doc, f, ensure_ascii=False, indent=2)
        self.log(f"📄 Распарсенный документ: {name}")

    def log_rule_prompt(self, rule_index: int, system_prompt: str, user_prompt: str):
        filepath = self.prompts_dir / f"rule_{rule_index:02d}_prompt.txt"
        with open(filepath, 'w', encoding='utf-8') as f:
            f.write(f"{'='*80}\n")
            f.write(f"RULE #{rule_index} - PROMPT\n")
            f.write(f"{'='*80}\n\n")
            f.write(f"--- SYSTEM PROMPT ---\n")
            f.write(system_prompt)
            f.write(f"\n\n--- USER PROMPT ---\n")
            f.write(user_prompt)

    def log_rule_response(self, rule_index: int, raw_response: str, parsed_violations: List[Dict]):
        filepath = self.responses_dir / f"rule_{rule_index:02d}_response.txt"
        with open(filepath, 'w', encoding='utf-8') as f:
            f.write(f"{'='*80}\n")
            f.write(f"RULE #{rule_index} - LLM RESPONSE\n")
            f.write(f"{'='*80}\n\n")
            f.write(f"--- RAW RESPONSE ---\n")
            f.write(raw_response)
            f.write(f"\n\n--- PARSED VIOLATIONS ---\n")
            f.write(json.dumps(parsed_violations, ensure_ascii=False, indent=2))

    def log_non_llm_result(self, rule_index: int, violations: List[Dict]):
        filepath = self.responses_dir / f"rule_{rule_index:02d}_non_llm.txt"
        with open(filepath, 'w', encoding='utf-8') as f:
            f.write(f"{'='*80}\n")
            f.write(f"RULE #{rule_index} - NON-LLM CHECK\n")
            f.write(f"{'='*80}\n\n")
            f.write(json.dumps(violations, ensure_ascii=False, indent=2))

    def log_final_results(self, violations: List[Dict]):
        filepath = self.log_dir / "final_results.json"
        with open(filepath, 'w', encoding='utf-8') as f:
            json.dump(violations, f, ensure_ascii=False, indent=2)
        self.log(f"📊 Финальные результаты: {len(violations)} нарушений")

    def log_rules_summary(self, rules_summary: List[Dict]):
        filepath = self.log_dir / "rules_summary.json"
        with open(filepath, 'w', encoding='utf-8') as f:
            json.dump(rules_summary, f, ensure_ascii=False, indent=2)

    def log_error(self, rule_index: int, error: str):
        filepath = self.responses_dir / f"rule_{rule_index:02d}_error.txt"
        with open(filepath, 'w', encoding='utf-8') as f:
            f.write(f"ERROR for rule #{rule_index}:\n{error}")
        self.log(f"❌ ОШИБКА правило #{rule_index}: {error}")


# Глобальный логгер
logger: PipelineLogger = None


# ============================================================================
# СИСТЕМНЫЙ ПРОМПТ
# ============================================================================

SYSTEM_PROMPT = """Ты — строгий аудитор документов.
Ты проверяешь ОДНО правило за один запрос.

Тебе даётся:
- фрагменты целевого документа (DOCX_*, XLSX_*);
- фрагменты шаблона (TEMPLATE_*) — если требуется сверка;
- параметры правила: RULE_INDEX, COMPARE, SCOPE, RULE_TITLE;
- текстовые инструкции правила (RULE_INSTRUCTIONS).

COMPARE:
- template — сверяй чанк целевого DOCX с чанком шаблона (игнорируй плейсхолдеры: даты, ФИО, должности).
- target_only — используй только целевой документ (DOCX или XLSX).
- cross_file — сверяй данные между DOCX и XLSX файлами.

Формат ответа — ОДИН JSON-объект:
- Если нарушений нет:
  {"status": "ok"}
- Если есть нарушения:
  {
    "status": "fail",
    "rule_index": <номер>,
    "rule_title": "<заголовок>",
    "нарушения": [
      {"Целевой документ": "...", "Различие": "..."}
    ]
  }

Что писать в полях:
- «Целевой документ» — что ФАКТИЧЕСКИ есть в документе (цитата) или «отсутствует»
- «Различие» — что ДОЛЖНО быть или в чём проблема

ВАЖНО:
- Верни ТОЛЬКО JSON, без пояснений.
- Одно правило = один объект ответа.
- Все нарушения по правилу собери в массив "нарушения".
"""


# ============================================================================
# МОДЕЛИ ДАННЫХ
# ============================================================================

@dataclass
class RuleSpec:
    """Спецификация правила проверки."""
    index: int
    title: str
    file: str  # "docx", "xlsx", "both"
    scope: str | List[str]
    compare: str  # "template", "target_only", "cross_file"
    llm: bool
    instructions: List[str]
    expected: Optional[str] = None


def load_rules(rules_path: Path) -> List[RuleSpec]:
    """Загружает правила из JSON файла."""
    with open(rules_path, 'r', encoding='utf-8') as f:
        data = json.load(f)

    rules = []
    for r in data.get("правила", []):
        rules.append(RuleSpec(
            index=r["index"],
            title=r["title"],
            file=r.get("file", "docx"),
            scope=r["scope"],
            compare=r["compare"],
            llm=r.get("llm", True),
            instructions=r.get("content", []),
            expected=r.get("expected")
        ))
    return rules


# ============================================================================
# ПОСТРОЕНИЕ КОНТЕКСТА
# ============================================================================

def build_context(spec: RuleSpec, docx_doc: dict, xlsx_doc: dict, template_doc: dict) -> str:
    """Формирует контекст для LLM."""
    parts = []
    scopes = [spec.scope] if isinstance(spec.scope, str) else spec.scope

    if spec.compare == "template":
        for scope in scopes:
            docx_content = docx_doc.get(scope, "(чанк отсутствует)")
            template_content = template_doc.get(scope, "(чанк отсутствует)")

            # Препроцессинг для правила #3
            if spec.index == 3 and scope == "текст_приказа":
                docx_content = normalize_text_for_rule3(docx_content)
                template_content = normalize_text_for_rule3(template_content)

            parts.append(f"[DOCX_{scope}]")
            parts.append(docx_content)
            parts.append(f"[/DOCX_{scope}]")

            parts.append(f"[TEMPLATE_{scope}]")
            parts.append(template_content)
            parts.append(f"[/TEMPLATE_{scope}]")

    elif spec.compare == "target_only":
        for scope in scopes:
            if spec.file == "docx":
                parts.append(f"[DOCX_{scope}]")
                parts.append(docx_doc.get(scope, "(чанк отсутствует)"))
                parts.append(f"[/DOCX_{scope}]")
            elif spec.file == "xlsx":
                parts.append(f"[XLSX_{scope}]")
                parts.append(xlsx_doc.get(scope, "(чанк отсутствует)"))
                parts.append(f"[/XLSX_{scope}]")

    elif spec.compare == "cross_file":
        for scope in scopes:
            if scope.startswith("docx."):
                chunk_name = scope.replace("docx.", "")
                parts.append(f"[DOCX_{chunk_name}]")
                parts.append(docx_doc.get(chunk_name, "(чанк отсутствует)"))
                parts.append(f"[/DOCX_{chunk_name}]")
            elif scope.startswith("xlsx."):
                chunk_name = scope.replace("xlsx.", "")
                parts.append(f"[XLSX_{chunk_name}]")
                parts.append(xlsx_doc.get(chunk_name, "(чанк отсутствует)"))
                parts.append(f"[/XLSX_{chunk_name}]")

    return "\n".join(parts)


def build_user_prompt(spec: RuleSpec, context: str) -> str:
    """Формирует пользовательский промпт для LLM."""
    scope_str = spec.scope if isinstance(spec.scope, str) else ", ".join(spec.scope)
    instructions = "\n".join(f"- {instr}" for instr in spec.instructions)

    return f"""RULE_INDEX: {spec.index}
COMPARE: {spec.compare}
FILE: {spec.file}
SCOPE: {scope_str}
RULE_TITLE: {spec.title}
RULE_INSTRUCTIONS:
{instructions}

CONTEXT:
{context}
"""


# ============================================================================
# LLM ВЫЗОВЫ
# ============================================================================

def call_llm(system_prompt: str, user_prompt: str, model: str) -> str:
    """Вызывает LLM через OpenAI API."""
    client = OpenAI()
    response = client.chat.completions.create(
        model=model,
        messages=[
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_prompt}
        ],
        temperature=0.0
    )
    return response.choices[0].message.content


def parse_llm_response(response: str, spec: RuleSpec) -> List[dict]:
    """Парсит JSON ответ от LLM."""
    text = response.strip()

    if text.startswith("```"):
        lines = text.split("\n")[1:]
        if lines and lines[-1].strip() == "```":
            lines = lines[:-1]
        text = "\n".join(lines)

    try:
        parsed = json.loads(text)

        if parsed.get("status") == "ok":
            return []

        violations = parsed.get("нарушения", [])
        for v in violations:
            v["rule_index"] = spec.index
            v["rule_title"] = spec.title

        return violations

    except json.JSONDecodeError as e:
        return [{
            "rule_index": spec.index,
            "rule_title": spec.title,
            "Целевой документ": "(ошибка парсинга)",
            "Различие": f"LLM вернул невалидный JSON: {e}"
        }]


# ============================================================================
# ПРОВЕРКИ
# ============================================================================

def check_filename_rule(spec: RuleSpec, docx_doc: dict) -> List[dict]:
    """Проверка имени файла без LLM (правило #2)."""
    expected = spec.expected or "3.6 Приказ о проведении выхода"
    actual = docx_doc.get("имя_файла", "")
    actual_name = actual.replace(".docx", "").replace(".DOCX", "")

    if expected.lower() in actual_name.lower():
        return []

    return [{
        "rule_index": spec.index,
        "rule_title": spec.title,
        "Целевой документ": actual,
        "Различие": f"Ожидалось название, содержащее: '{expected}'"
    }]


def run_llm_check(spec: RuleSpec, docx_doc: dict, xlsx_doc: dict,
                  template_doc: dict, model: str) -> List[dict]:
    """Выполняет проверку одного правила через LLM."""
    global logger

    context = build_context(spec, docx_doc, xlsx_doc, template_doc)
    user_prompt = build_user_prompt(spec, context)

    # Логируем промпт
    if logger:
        logger.log_rule_prompt(spec.index, SYSTEM_PROMPT, user_prompt)

    response = call_llm(SYSTEM_PROMPT, user_prompt, model)
    violations = parse_llm_response(response, spec)

    # Логируем ответ
    if logger:
        logger.log_rule_response(spec.index, response, violations)

    return violations


def run_audit(rules: List[RuleSpec], docx_doc: dict, xlsx_doc: dict,
              template_doc: dict, model: str, print_prompts: bool = False) -> List[dict]:
    """Запускает проверку всех правил."""
    global logger
    all_violations = []
    rules_summary = []

    # Сначала правила без LLM
    for spec in rules:
        if not spec.llm:
            if spec.index == 2:
                violations = check_filename_rule(spec, docx_doc)
                all_violations.extend(violations)
                if logger:
                    logger.log_non_llm_result(spec.index, violations)
                status = "ok" if not violations else "fail"
                rules_summary.append({"index": spec.index, "title": spec.title, "status": status})
                icon = "✅" if not violations else "❌"
                print(f"  {icon} Правило #{spec.index}: {spec.title}")

    # Режим отладки
    if print_prompts:
        for spec in [r for r in rules if r.llm]:
            context = build_context(spec, docx_doc, xlsx_doc, template_doc)
            user_prompt = build_user_prompt(spec, context)
            print(f"\n{'='*70}")
            print(f"ПРАВИЛО #{spec.index}: {spec.title}")
            print(f"{'='*70}")
            print(user_prompt)
        return all_violations

    # Параллельный запуск LLM
    llm_rules = [r for r in rules if r.llm]

    with ThreadPoolExecutor(max_workers=MAX_WORKERS) as executor:
        futures = {
            executor.submit(run_llm_check, spec, docx_doc, xlsx_doc, template_doc, model): spec
            for spec in llm_rules
        }

        for future in as_completed(futures):
            spec = futures[future]
            try:
                violations = future.result()
                all_violations.extend(violations)
                status = "ok" if not violations else "fail"
                rules_summary.append({"index": spec.index, "title": spec.title, "status": status})
                icon = "✅" if not violations else "❌"
                print(f"  {icon} Правило #{spec.index}: {spec.title}")
            except Exception as e:
                if logger:
                    logger.log_error(spec.index, str(e))
                rules_summary.append({"index": spec.index, "title": spec.title, "status": "error"})
                all_violations.append({
                    "rule_index": spec.index,
                    "rule_title": spec.title,
                    "Целевой документ": "(ошибка)",
                    "Различие": str(e)
                })
                print(f"  ⚠️  Правило #{spec.index}: ошибка - {e}")

    # Сохраняем сводку по правилам
    if logger:
        rules_summary.sort(key=lambda x: x["index"])
        logger.log_rules_summary(rules_summary)

    all_violations.sort(key=lambda x: x.get("rule_index", 0))
    return all_violations


def save_results(violations: List[dict], output_path: Path):
    """Сохраняет результаты в Excel файл."""
    if not violations:
        df = pd.DataFrame(columns=["rule_index", "rule_title", "Целевой документ", "Различие"])
    else:
        df = pd.DataFrame(violations)

    df.to_excel(output_path, index=False)


# ============================================================================
# MAIN
# ============================================================================

def main():
    global logger

    parser = argparse.ArgumentParser(description="LLM-аудит Приказа о проведении выхода")
    parser.add_argument("--docx", required=True, help="Путь к целевому DOCX файлу")
    parser.add_argument("--xlsx", required=True, help="Путь к целевому XLSX файлу")
    parser.add_argument("--template", required=True, help="Путь к шаблону DOCX")
    parser.add_argument("--rules", default=str(DEFAULT_RULES), help="Путь к JSON с правилами")
    parser.add_argument("--model", default=DEFAULT_MODEL, help=f"Модель OpenAI (default: {DEFAULT_MODEL})")
    parser.add_argument("--out-xlsx", help="Путь для сохранения результатов в Excel")
    parser.add_argument("--print-prompts", action="store_true", help="Показать промпты без вызова LLM")

    args = parser.parse_args()

    # Проверяем существование файлов
    for path, name in [(args.docx, "DOCX"), (args.xlsx, "XLSX"), (args.template, "Шаблон")]:
        if not Path(path).exists():
            print(f"❌ Файл не найден: {path}")
            sys.exit(1)

    # Создаём сессию логирования
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    session_dir = PROJECT_DIR / "logs_result" / f"session_{timestamp}"
    logger = PipelineLogger(session_dir)

    print(f"📄 DOCX:    {Path(args.docx).name}")
    print(f"📊 XLSX:    {Path(args.xlsx).name}")
    print(f"📋 Шаблон:  {Path(args.template).name}")
    print(f"🤖 Модель:  {args.model}")
    print(f"📁 Сессия:  {session_dir.name}")
    print()

    logger.log(f"DOCX: {args.docx}")
    logger.log(f"XLSX: {args.xlsx}")
    logger.log(f"Template: {args.template}")
    logger.log(f"Model: {args.model}")

    # Парсим документы
    print("⏳ Парсинг документов...")
    docx_doc = parse_prikaz_vyhod(args.docx)
    xlsx_doc = parse_grafik_obhod(args.xlsx)
    template_doc = parse_prikaz_vyhod(args.template)

    logger.log_parsed_doc(docx_doc, "target_docx")
    logger.log_parsed_doc(xlsx_doc, "target_xlsx")
    logger.log_parsed_doc(template_doc, "template_docx")

    print("   ✅ Документы загружены")
    print()

    # Загружаем правила
    rules = load_rules(Path(args.rules))
    logger.log(f"Загружено правил: {len(rules)}")
    print(f"📋 Загружено правил: {len(rules)}")
    print()

    # Запускаем аудит
    print("🔍 Проверка правил:")
    violations = run_audit(rules, docx_doc, xlsx_doc, template_doc, args.model, args.print_prompts)

    # Сохраняем результаты
    logger.log_final_results(violations)

    # Excel отчёт
    excel_path = args.out_xlsx if args.out_xlsx else session_dir / "audit_result.xlsx"
    save_results(violations, Path(excel_path))
    logger.log(f"Excel отчёт: {excel_path}")

    # Выводим итоги
    print()
    print("=" * 60)
    if violations:
        print(f"❌ Найдено нарушений: {len(violations)}")
        for v in violations:
            print(f"\n  Правило #{v['rule_index']}: {v['rule_title']}")
            print(f"    Целевой документ: {v.get('Целевой документ', '?')}")
            print(f"    Различие: {v.get('Различие', '?')}")
    else:
        print("✅ Нарушений не найдено")

    print(f"\n📊 Результаты сохранены: {session_dir}")

    sys.exit(1 if violations else 0)


if __name__ == "__main__":
    main()
