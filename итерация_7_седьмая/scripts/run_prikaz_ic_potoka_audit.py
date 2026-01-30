#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Скрипт LLM-аудита документа "Приказ о создании ИЦ потока".

Выполняет автоматическую проверку документа по правилам ТЗ с использованием LLM.
Поддерживает два варианта: бумажный (10 правил) и электронный (11 правил).

Типы проверок:
- target_only: проверка только целевого документа
- template: сверка текста целевого документа с шаблоном
- cross_check: сверка сущностей между чанками целевого документа

Все артефакты сессии сохраняются в директорию:
logs_result/session_YYYYMMDD_HHMMSS/
├── prompts/           # Промпты для каждого правила
├── responses/         # Ответы LLM
├── parsed_docs/       # Распарсенные документы
├── pipeline.log       # Лог выполнения
├── final_results.json # JSON с нарушениями
├── rules_summary.json # Сводка по правилам
└── audit_result.xlsx  # Excel-отчёт
"""

import argparse
import json
import os
import sys
import re
from concurrent.futures import ThreadPoolExecutor, as_completed
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Union

import pandas as pd
from openai import OpenAI

# Добавляем путь к парсеру
sys.path.insert(0, str(Path(__file__).parent))
from parser_prikaz_ic_potoka import parse_prikaz_ic_potoka

# Добавляем путь к итерация_vision модулям
sys.path.insert(0, str(Path(__file__).parent.parent.parent))
# Vision Parser будет импортирован при необходимости


# ============================================================================
# ПРЕПРОЦЕССИНГ ДЛЯ ПРАВИЛА #3 (сверка текста с шаблоном)
# ============================================================================

def normalize_text_for_rule3(text: str) -> str:
    """
    Нормализует текст для правила #3, заменяя плейсхолдеры на унифицированные метки.
    Применяется и к целевому документу и к шаблону.
    """
    # 0. Убираем лишние переносы строк и нормализуем пробелы
    # Заменяем множественные пробелы на один
    text = re.sub(r'[ \t]+', ' ', text)
    # Заменяем перенос строки + пробелы на один перенос
    text = re.sub(r'\n\s*', '\n', text)

    # 1. Нормализуем даты: __.__.202_, конкретные даты → [ДАТА]
    text = re.sub(r'\d{1,2}\.\d{1,2}\.\d{4}', '[ДАТА]', text)
    text = re.sub(r'_{2,}\.\s*_{2,}\.\s*202_?', '[ДАТА]', text)

    # 2. Нормализуем п.4: должность+ФИО → [ДОЛЖНОСТЬ_ФИО]
    # Шаблон: "(Указать наименование должности) организовать..."
    # Целевой: "Начальнику службы гостиничного хозяйства Ивановой А.А. организовать..."
    text = re.sub(
        r'\(Указать наименование должности\)',
        '[ДОЛЖНОСТЬ_ФИО]',
        text
    )
    # Заменяем конкретную должность+ФИО перед "организовать"
    text = re.sub(
        r'[А-ЯЁа-яё\s]+[А-ЯЁ][а-яё]+\s+[А-ЯЁ]\.[А-ЯЁ]\.\s+организовать',
        '[ДОЛЖНОСТЬ_ФИО] организовать',
        text
    )

    # 3. Нормализуем п.3: принудительно добавляем перенос между [ДАТА] и [ДОЛЖНОСТЬ_ФИО]
    # Это нужно потому что в документе пункты могут быть слиты в одну строку
    text = re.sub(
        r'\[ДАТА\]\s*\[ДОЛЖНОСТЬ_ФИО\]',
        r'[ДАТА]\n[ДОЛЖНОСТЬ_ФИО]',
        text
    )

    # 4. Нормализуем подписанта
    # Шаблон: "Генеральный директор                    И.О. Фамилия"
    # Целевой: "Генеральный директор - должность        ФИО подписанта - А.В. Петров"
    lines = text.split('\n')
    for i in range(len(lines) - 1, -1, -1):
        line = lines[i].strip()
        if line and ('директор' in line.lower() or 'должность' in line.lower() or
                     'фамилия' in line.lower() or 'фио' in line.lower() or
                     re.search(r'[А-ЯЁ]\.[А-ЯЁ]\.', line)):
            lines[i] = '[ПОДПИСАНТ]'
            break
    text = '\n'.join(lines)

    return text


# ============================================================================
# ЛОГИРОВАНИЕ
# ============================================================================

class PipelineLogger:
    """Логгер для всего пайплайна аудита."""

    def __init__(self, log_dir: Path):
        """Инициализация логгера."""
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
        """Инициализирует главный лог-файл."""
        with open(self.main_log, 'w', encoding='utf-8') as f:
            f.write(f"{'='*80}\n")
            f.write(f"AUDIT PIPELINE LOG - Приказ о создании ИЦ потока\n")
            f.write(f"Started: {datetime.now().isoformat()}\n")
            f.write(f"{'='*80}\n\n")

    def log(self, message: str):
        """Записывает сообщение в главный лог."""
        timestamp = datetime.now().strftime("%H:%M:%S")
        with open(self.main_log, 'a', encoding='utf-8') as f:
            f.write(f"[{timestamp}] {message}\n")
        print(f"[{timestamp}] {message}", file=sys.stderr)

    def log_parsed_doc(self, doc: Dict[str, Any], name: str):
        """Сохраняет распарсенный документ."""
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
        """Логирует ошибку."""
        filepath = self.responses_dir / f"rule_{rule_index:02d}_error.txt"
        with open(filepath, 'w', encoding='utf-8') as f:
            f.write(f"ERROR for rule #{rule_index}:\n{error}")
        self.log(f"❌ ОШИБКА для правила #{rule_index}: {error}")


# Глобальный логгер
logger: PipelineLogger = None


# ============================================================================
# КОНСТАНТЫ И ПРОМПТЫ
# ============================================================================

# Модель по умолчанию
DEFAULT_MODEL = "gpt-4.1-mini"

# Количество параллельных запросов к LLM
# ВАЖНО: gpt-4.1-mini нестабилен при параллельных запросах, отключено
MAX_WORKERS = 1

# Системный промпт
SYSTEM_PROMPT = """Ты — строгий аудитор документов.
Ты проверяешь ОДНО правило за один запрос.

Тебе даётся:
- фрагменты целевого документа (TARGET_*);
- фрагменты шаблона (TEMPLATE_*) — если требуется сверка;
- параметры правила: RULE_INDEX, COMPARE, SCOPE, RULE_TITLE;
- текстовые инструкции правила (RULE_INSTRUCTIONS).

COMPARE:
- template — сверяй чанк целевого с чанком шаблона (игнорируй плейсхолдеры: даты, ФИО, должности).
- target_only — используй только целевой документ.
- cross_check — сверяй два чанка целевого документа между собой.

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
# DATACLASS ДЛЯ ПРАВИЛ
# ============================================================================

@dataclass
class RuleSpec:
    """Спецификация одного правила проверки."""
    index: int
    title: str
    scope: Union[str, List[str]]
    compare: str
    llm: bool
    instructions: List[str] = field(default_factory=list)
    context_filter: Dict[str, List[str]] = field(default_factory=dict)
    context_filter_mode: str = "paragraphs"  # "paragraphs" или "headers_only"


# ============================================================================
# ФИЛЬТРАЦИЯ КОНТЕКСТА
# ============================================================================

def extract_matching_paragraphs(text: str, patterns: List[str], headers_only: bool = False) -> str:
    """
    Извлекает из текста только строки/пункты, соответствующие паттернам.

    Args:
        text: Исходный текст
        patterns: Список regex паттернов для начала строки
        headers_only: Если True, извлекает только совпадающие строки без продолжения

    Примеры:
      patterns = ["^2\\.1", "^2\\.2"], headers_only=False
      text = "1. Общие\n2.1. Ответственный...\nПродолжение\n2.2. И.о....\n3. Порядок"
      Результат = "2.1. Ответственный...\nПродолжение\n2.2. И.о...."

      patterns = ["^1\\. ", "^2\\. "], headers_only=True
      text = "1. Общие положения\nТекст...\n2. Порядок\nТекст2..."
      Результат = "1. Общие положения\n2. Порядок"
    """
    if not patterns:
        return text

    lines = text.split('\n')
    result_lines = []

    # Компилируем паттерны
    compiled_patterns = [re.compile(p) for p in patterns]

    if headers_only:
        # Режим только заголовков — извлекаем только совпадающие строки
        for line in lines:
            stripped = line.strip()
            if any(p.match(stripped) for p in compiled_patterns):
                result_lines.append(stripped)
    else:
        # Режим полных абзацев — извлекаем пункт с продолжением
        capturing = False
        # Паттерн для определения начала нового пункта (любой пункт типа "X." или "X.Y")
        new_section_pattern = re.compile(r'^(\d+\.|\d+\.\d+\.?)\s')

        for line in lines:
            stripped = line.strip()
            matches_our_pattern = any(p.match(stripped) for p in compiled_patterns)

            if matches_our_pattern:
                capturing = True
                result_lines.append(line)
            elif capturing:
                if new_section_pattern.match(stripped) and not matches_our_pattern:
                    capturing = False
                else:
                    result_lines.append(line)

    filtered_text = '\n'.join(result_lines).strip()

    if not filtered_text:
        return f"[Фильтр: не найдено пунктов по паттернам {patterns}]"

    return filtered_text


# ============================================================================
# ЗАГРУЗКА ПРАВИЛ
# ============================================================================

def load_rules(rules_path: str) -> List[RuleSpec]:
    """Загружает правила из JSON-файла."""
    with open(rules_path, 'r', encoding='utf-8') as f:
        data = json.load(f)

    rules = []
    for rule in data.get("правила", []):
        spec = RuleSpec(
            index=rule["index"],
            title=rule["title"],
            scope=rule["scope"],
            compare=rule["compare"],
            llm=rule["llm"],
            instructions=rule.get("content", []),
            context_filter=rule.get("context_filter", {}),
            context_filter_mode=rule.get("context_filter_mode", "paragraphs")
        )
        rules.append(spec)

    return rules


# ============================================================================
# ПОСТРОЕНИЕ КОНТЕКСТА ДЛЯ LLM
# ============================================================================

def build_context_for_rule(
    spec: RuleSpec,
    target_doc: Dict[str, Any],
    template_doc: Dict[str, Any]
) -> str:
    """Строит контекст для правила в зависимости от типа проверки."""
    context_parts = []
    scopes = [spec.scope] if isinstance(spec.scope, str) else spec.scope

    # Функция для применения фильтра контекста
    def apply_filter(content: str, scope: str) -> str:
        """Применяет context_filter если задан для данного scope."""
        if spec.context_filter and scope in spec.context_filter:
            patterns = spec.context_filter[scope]
            headers_only = spec.context_filter_mode == "headers_only"
            return extract_matching_paragraphs(content, patterns, headers_only)
        return content

    if spec.compare == "target_only":
        # Только чанки целевого документа
        for scope in scopes:
            content = target_doc.get(scope, "")
            content = apply_filter(content, scope)
            context_parts.append(f"[TARGET_{scope}]")
            context_parts.append(content)
            context_parts.append(f"[/TARGET_{scope}]")

    elif spec.compare == "template":
        # Чанки целевого + шаблона
        for scope in scopes:
            target_content = target_doc.get(scope, "")
            template_content = template_doc.get(scope, "")

            # Препроцессинг для правила #3
            if spec.index == 3 and scope == "текст_приказа":
                target_content = normalize_text_for_rule3(target_content)
                template_content = normalize_text_for_rule3(template_content)

            # Применяем фильтр контекста
            target_content = apply_filter(target_content, scope)
            template_content = apply_filter(template_content, scope)

            context_parts.append(f"[TARGET_{scope}]")
            context_parts.append(target_content)
            context_parts.append(f"[/TARGET_{scope}]")

            context_parts.append(f"[TEMPLATE_{scope}]")
            context_parts.append(template_content)
            context_parts.append(f"[/TEMPLATE_{scope}]")

    elif spec.compare == "cross_check":
        # Несколько чанков целевого документа для сравнения
        for scope in scopes:
            content = target_doc.get(scope, "")
            content = apply_filter(content, scope)
            context_parts.append(f"[TARGET_{scope}]")
            context_parts.append(content)
            context_parts.append(f"[/TARGET_{scope}]")

    return "\n".join(context_parts)


def build_user_prompt(
    spec: RuleSpec,
    target_doc: Dict[str, Any],
    template_doc: Dict[str, Any]
) -> str:
    """Формирует пользовательский промпт для LLM."""
    scope_str = spec.scope if isinstance(spec.scope, str) else ", ".join(spec.scope)
    instructions_text = "\n".join(f"- {instr}" for instr in spec.instructions)
    context = build_context_for_rule(spec, target_doc, template_doc)

    prompt = f"""RULE_INDEX: {spec.index}
COMPARE: {spec.compare}
SCOPE: {scope_str}
RULE_TITLE: {spec.title}
RULE_INSTRUCTIONS:
{instructions_text}

CONTEXT:
{context}
"""
    return prompt


# ============================================================================
# NON-LLM ПРОВЕРКИ
# ============================================================================

def check_filename_rule(target_doc: Dict[str, Any], variant: str) -> List[Dict[str, Any]]:
    """Правило #2: проверка имени файла без LLM."""
    if variant == "electronic":
        expected_pattern = "1.3 Приказ о создании ИЦ потока эл"
    else:
        expected_pattern = "1.3 Приказ о создании ИЦ потока"

    actual = target_doc.get("имя_файла", "")
    actual_clean = actual.replace(".docx", "")

    # Проверяем вхождение паттерна
    if expected_pattern.lower() in actual_clean.lower():
        return []

    return [{
        "rule_index": 2,
        "rule_title": "Проверка имени файла документа.",
        "Целевой документ": actual_clean,
        "Различие": f"Ожидалось имя файла, содержащее: «{expected_pattern}»"
    }]


def check_non_llm_rule(
    spec: RuleSpec,
    target_doc: Dict[str, Any],
    variant: str = "paper"
) -> List[Dict[str, Any]]:
    """Выполняет проверку правила без LLM."""
    if spec.index == 2:
        return check_filename_rule(target_doc, variant)
    return []


# ============================================================================
# LLM ПРОВЕРКИ
# ============================================================================

def parse_json_response(raw_response: str, spec_index: int, spec_title: str) -> List[Dict[str, Any]]:
    """Парсит JSON-ответ от LLM."""
    text = raw_response.strip()

    # Убираем markdown код-блок
    if text.startswith("```"):
        lines = text.split("\n")
        if lines[0].startswith("```"):
            lines = lines[1:]
        if lines and lines[-1].strip() == "```":
            lines = lines[:-1]
        text = "\n".join(lines)

    # Санитизация JSON
    def sanitize_json_string(s: str) -> str:
        result_chars = []
        in_string = False
        escape_next = False

        for char in s:
            if escape_next:
                result_chars.append(char)
                escape_next = False
                continue

            if char == '\\' and in_string:
                result_chars.append(char)
                escape_next = True
                continue

            if char == '"':
                in_string = not in_string
                result_chars.append(char)
                continue

            if in_string and char in '\n\r\t':
                if char == '\n':
                    result_chars.append('\\n')
                elif char == '\r':
                    result_chars.append('\\r')
                elif char == '\t':
                    result_chars.append('\\t')
            else:
                result_chars.append(char)

        return ''.join(result_chars)

    text = sanitize_json_string(text)

    try:
        result = json.loads(text)

        if isinstance(result, dict):
            if result.get("status") == "ok":
                return []

            violations = result.get("нарушения", [])
            for v in violations:
                v["rule_index"] = result.get("rule_index", spec_index)
                v["rule_title"] = result.get("rule_title", spec_title)
            return violations

        if isinstance(result, list):
            return result

        return []

    except json.JSONDecodeError as e:
        print(f"[WARN] Не удалось распарсить JSON: {e}", file=sys.stderr)
        print(f"[WARN] Ответ: {text[:200]}...", file=sys.stderr)
        return []


def call_llm(
    messages: List[Dict[str, str]],
    model: str,
    temperature: float = 0.0
) -> str:
    """Вызывает LLM через OpenAI API."""
    client = OpenAI()

    response = client.chat.completions.create(
        model=model,
        messages=messages,
        temperature=temperature
    )

    return response.choices[0].message.content


def run_llm_rule_check(
    spec: RuleSpec,
    target_doc: Dict[str, Any],
    template_doc: Dict[str, Any],
    model: str,
    temperature: float = 0.0
) -> List[Dict[str, Any]]:
    """Проверка одного правила через LLM."""
    global logger

    user_prompt = build_user_prompt(spec, target_doc, template_doc)

    messages = [
        {"role": "system", "content": SYSTEM_PROMPT},
        {"role": "user", "content": user_prompt}
    ]

    if logger:
        logger.log_rule_prompt(spec.index, SYSTEM_PROMPT, user_prompt)

    raw_response = call_llm(messages, model, temperature)

    violations = parse_json_response(raw_response, spec.index, spec.title)

    for obj in violations:
        obj.setdefault("rule_index", spec.index)
        obj.setdefault("rule_title", spec.title)

    if logger:
        logger.log_rule_response(spec.index, raw_response, violations)

    return violations


# ============================================================================
# ОРКЕСТРАЦИЯ
# ============================================================================

def run_all_checks(
    rules: List[RuleSpec],
    target_doc: Dict[str, Any],
    template_doc: Dict[str, Any],
    model: str,
    variant: str = "paper",
    temperature: float = 0.0,
    print_prompts: bool = False
) -> List[Dict[str, Any]]:
    """Выполняет все проверки."""
    global logger
    all_violations = []

    # Non-LLM правила
    for spec in rules:
        if not spec.llm:
            if logger:
                logger.log(f"🔧 Запуск non-LLM правила #{spec.index}: {spec.title}")
            violations = check_non_llm_rule(spec, target_doc, variant)
            all_violations.extend(violations)

            if logger:
                logger.log_non_llm_result(spec.index, violations)

            if print_prompts:
                print(f"\n{'='*60}")
                print(f"[NON-LLM] Правило #{spec.index}: {spec.title}")
                print(f"Результат: {violations}")

    # LLM правила
    llm_rules = [r for r in rules if r.llm]

    if print_prompts:
        for spec in llm_rules:
            user_prompt = build_user_prompt(spec, target_doc, template_doc)
            print(f"\n{'='*60}")
            print(f"[LLM] Правило #{spec.index}: {spec.title}")
            print(f"{'='*60}")
            print("\n--- SYSTEM PROMPT ---")
            print(SYSTEM_PROMPT)
            print("\n--- USER PROMPT ---")
            print(user_prompt)

            if logger:
                logger.log_rule_prompt(spec.index, SYSTEM_PROMPT, user_prompt)
        return all_violations

    # Параллельное выполнение
    if logger:
        logger.log(f"🚀 Запуск {len(llm_rules)} LLM-проверок (max_workers={MAX_WORKERS})")

    with ThreadPoolExecutor(max_workers=MAX_WORKERS) as executor:
        futures = {
            executor.submit(
                run_llm_rule_check,
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
                if logger:
                    logger.log(f"✅ Правило #{spec.index} проверено, нарушений: {len(violations)}")
            except Exception as e:
                error_msg = str(e)
                if logger:
                    logger.log_error(spec.index, error_msg)
                print(f"[ERROR] Правило #{spec.index}: {e}", file=sys.stderr)

    return sorted(all_violations, key=lambda x: x.get("rule_index", 0))


# ============================================================================
# СОХРАНЕНИЕ РЕЗУЛЬТАТОВ
# ============================================================================

def save_to_excel(violations: List[Dict[str, Any]], output_path: str) -> None:
    """Сохраняет результаты в Excel."""
    if not violations:
        df = pd.DataFrame(columns=[
            "rule_index",
            "rule_title",
            "Целевой документ",
            "Различие"
        ])
    else:
        df = pd.DataFrame(violations)
        columns_order = ["rule_index", "rule_title", "Целевой документ", "Различие"]
        existing_cols = [c for c in columns_order if c in df.columns]
        extra_cols = [c for c in df.columns if c not in columns_order]
        df = df[existing_cols + extra_cols]

    Path(output_path).parent.mkdir(parents=True, exist_ok=True)
    df.to_excel(output_path, index=False, engine='openpyxl')
    print(f"✅ Excel сохранён: {output_path}", file=sys.stderr)


# ============================================================================
# MAIN
# ============================================================================

def main():
    """Точка входа CLI."""
    global logger

    parser = argparse.ArgumentParser(
        description="LLM-аудит документа 'Приказ о создании ИЦ потока'"
    )
    parser.add_argument(
        "--target",
        required=True,
        help="Путь к целевому документу (.docx)"
    )
    parser.add_argument(
        "--template",
        required=True,
        help="Путь к шаблону (.docx)"
    )
    parser.add_argument(
        "--rules",
        default=None,
        help="Путь к JSON с правилами"
    )
    parser.add_argument(
        "--variant",
        choices=["paper", "electronic"],
        default="paper",
        help="Вариант документа: paper (бумажный) или electronic (электронный)"
    )
    parser.add_argument(
        "--model",
        default=DEFAULT_MODEL,
        help=f"Модель OpenAI (по умолчанию: {DEFAULT_MODEL})"
    )
    parser.add_argument(
        "--out-xlsx",
        default=None,
        help="Путь для сохранения Excel"
    )
    parser.add_argument(
        "--print-prompts",
        action="store_true",
        help="Режим отладки: только промпты без вызова LLM"
    )
    parser.add_argument(
        "--temperature",
        type=float,
        default=0.0,
        help="Температура генерации"
    )
    parser.add_argument(
        "--session-dir",
        default=None,
        help="Директория сессии"
    )
    parser.add_argument(
        "--use-vision",
        action="store_true",
        help="Использовать Vision Pipeline (gpt-4.1-mini) вместо python-docx парсера"
    )

    args = parser.parse_args()

    # Определяем путь к правилам
    script_dir = Path(__file__).parent.parent
    if args.rules:
        rules_path = args.rules
    else:
        if args.variant == "electronic":
            rules_path = script_dir / "config" / "tz_prikaz_ic_potoka_el.json"
        else:
            rules_path = script_dir / "config" / "tz_prikaz_ic_potoka.json"

    # Инициализируем директорию сессии
    if args.session_dir:
        session_dir = Path(args.session_dir)
    else:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        session_dir = script_dir / "logs_result" / f"session_{timestamp}"

    logger = PipelineLogger(session_dir)
    logger.log(f"🚀 Запуск аудита (вариант: {args.variant})")
    logger.log(f"   Сессия: {session_dir}")
    logger.log(f"   Целевой документ: {args.target}")
    logger.log(f"   Шаблон: {args.template}")
    logger.log(f"   Правила: {rules_path}")
    logger.log(f"   Модель: {args.model}")

    # Проверяем файлы
    if not Path(args.target).exists():
        logger.log(f"❌ Целевой документ не найден: {args.target}")
        sys.exit(1)

    if not Path(args.template).exists():
        logger.log(f"❌ Шаблон не найден: {args.template}")
        sys.exit(1)

    if not Path(rules_path).exists():
        logger.log(f"❌ Файл правил не найден: {rules_path}")
        sys.exit(1)

    # Парсим документы
    if args.use_vision:
        # Vision Pipeline через gpt-4.1-mini
        logger.log(f"🔮 Использование Vision Pipeline для парсинга...")
        try:
            from итерация_vision.vision_parser import VisionParser
        except ImportError as e:
            logger.log(f"❌ Не удалось импортировать VisionParser: {e}")
            logger.log(f"   Убедитесь, что установлены зависимости: pip install pdf2image tenacity Pillow")
            sys.exit(1)

        # Путь к конфигурации чанков для Vision
        vision_config_path = script_dir / "config" / "chunks_vision.json"
        if not vision_config_path.exists():
            logger.log(f"❌ Конфигурация Vision не найдена: {vision_config_path}")
            sys.exit(1)

        logger.log(f"   Vision конфиг: {vision_config_path}")

        # Парсим целевой документ через Vision
        # Передаём session_dir для логирования промптов Vision
        logger.log(f"📄 Vision-парсинг целевого документа...")
        vision_parser_target = VisionParser(
            str(vision_config_path),
            log_dir=str(session_dir / "vision_target")
        )
        target_doc = vision_parser_target.parse(args.target)
        logger.log_parsed_doc(target_doc, "target_doc_vision")

        # Парсим шаблон через Vision
        logger.log(f"📄 Vision-парсинг шаблона...")
        vision_parser_template = VisionParser(
            str(vision_config_path),
            log_dir=str(session_dir / "vision_template")
        )
        template_doc = vision_parser_template.parse(args.template)
        logger.log_parsed_doc(template_doc, "template_doc_vision")
    else:
        # Legacy парсер через python-docx
        logger.log(f"📄 Парсинг целевого документа (legacy)...")
        target_doc = parse_prikaz_ic_potoka(args.target)
        logger.log_parsed_doc(target_doc, "target_doc")

        logger.log(f"📄 Парсинг шаблона (legacy)...")
        template_doc = parse_prikaz_ic_potoka(args.template)
        logger.log_parsed_doc(template_doc, "template_doc")

    # Загружаем правила
    logger.log(f"📋 Загрузка правил...")
    rules = load_rules(str(rules_path))
    logger.log(f"   Загружено {len(rules)} правил")

    # Сохраняем сводку правил
    rules_summary = [
        {"index": r.index, "title": r.title, "llm": r.llm, "compare": r.compare}
        for r in rules
    ]
    with open(session_dir / "rules_summary.json", 'w', encoding='utf-8') as f:
        json.dump(rules_summary, f, ensure_ascii=False, indent=2)

    # Выполняем проверки
    logger.log(f"🔍 Запуск проверок...")
    violations = run_all_checks(
        rules=rules,
        target_doc=target_doc,
        template_doc=template_doc,
        model=args.model,
        variant=args.variant,
        temperature=args.temperature,
        print_prompts=args.print_prompts
    )

    # Сохраняем результаты
    logger.log_final_results(violations)

    print(json.dumps(violations, ensure_ascii=False, indent=2))

    # Excel
    if args.out_xlsx:
        xlsx_path = args.out_xlsx
    else:
        xlsx_path = str(session_dir / "audit_result.xlsx")
    save_to_excel(violations, xlsx_path)
    logger.log(f"📊 Excel сохранён: {xlsx_path}")

    # Итоговая статистика
    logger.log(f"{'='*60}")
    logger.log(f"📊 Итого нарушений: {len(violations)}")
    if violations:
        by_rule = {}
        for v in violations:
            idx = v.get("rule_index", "?")
            by_rule.setdefault(idx, 0)
            by_rule[idx] += 1
        logger.log("   По правилам:")
        for idx in sorted(by_rule.keys()):
            logger.log(f"   - Правило #{idx}: {by_rule[idx]} нарушений")

    logger.log(f"✅ Аудит завершён. Сессия: {session_dir}")


if __name__ == "__main__":
    main()
