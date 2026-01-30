#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Скрипт LLM-аудита документа "Приказ о создании ИЦ".

Выполняет автоматическую проверку документа по правилам ТЗ с использованием LLM.
Правила загружаются из JSON-файла конфигурации.

Типы проверок:
- target_only: проверка только целевого документа (заполняемость плейсхолдеров)
- template: сверка текста целевого документа с шаблоном
- cross_check: сверка сущностей между чанками целевого документа

Все артефакты сессии (логи, промпты, ответы, Excel) сохраняются в единую директорию:
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
from concurrent.futures import ThreadPoolExecutor, as_completed
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Union

import pandas as pd
from openai import OpenAI

# Добавляем путь к парсеру и моделям
sys.path.insert(0, str(Path(__file__).parent))
from parser_prikaz_ic_docs import parse_prikaz_ic

import re

# ============================================================================
# ПРЕПРОЦЕССИНГ ДЛЯ ПРАВИЛА #3 (сверка текста с шаблоном)
# ============================================================================

def normalize_text_for_rule3(text: str) -> str:
    """
    Нормализует текст для правила #3, заменяя плейсхолдеры на унифицированные метки.
    Применяется и к целевому документу и к шаблону, чтобы они стали идентичны
    (если нет структурных различий).

    Плейсхолдеры:
    - Даты: __.__.202_, 27.08.2025, 27.0., 27.08. и т.д. → [ДАТА]
    - Должность+ФИО в п.4: текст до "организовать" → [ДОЛЖНОСТЬ_ФИО]
    - Подписант: последняя строка с должностью и ФИО → [ПОДПИСАНТ]
    """
    # 1. Нормализуем п.3: дата после "с" в конце строки
    # Шаблон: "с __.__.202_" → Целевой: "с 27.08.2025" или "с 27.0."
    text = re.sub(
        r'(приступить к заполнению разделов по своим показателям с\s*)[^\n]+',
        r'\1[ДАТА]',
        text
    )

    # 2. Нормализуем п.4: должность+ФИО до "организовать" и дата после "до"
    # Шаблон: "4. (Указать наименование должности) организовать..."
    # Целевой: "4. Начальнику отдела продаж Гуния И.Д. организовать..."
    text = re.sub(
        r'4\.\s*.+?\s+организовать',
        r'4. [ДОЛЖНОСТЬ_ФИО] организовать',
        text
    )
    # Дата после "в срок до"
    text = re.sub(
        r'(в срок до\s*)[^\n]+',
        r'\1[ДАТА]',
        text
    )

    # 3. Нормализуем подписанта (последняя строка с должностью и ФИО)
    # Шаблон: "Генеральный директор                    И.О. Фамилия"
    # Целевой: "фыв - должность                        ФИО подписанта - А.В. ыв"
    # Ищем строку с должностью и ФИО в конце текста
    lines = text.split('\n')
    for i in range(len(lines) - 1, -1, -1):
        line = lines[i].strip()
        # Если строка похожа на подписанта (содержит пробелы между частями)
        if line and ('директор' in line.lower() or 'должность' in line.lower()
                     or 'фамилия' in line.lower() or 'фио' in line.lower()
                     or re.search(r'\s{5,}', line)):  # много пробелов между частями
            lines[i] = '[ПОДПИСАНТ]'
            break
    text = '\n'.join(lines)

    return text


# ============================================================================
# ЛОГИРОВАНИЕ
# ============================================================================

class PipelineLogger:
    """
    Логгер для всего пайплайна аудита.
    Сохраняет все этапы в отдельные файлы для отладки.
    """

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
            f.write(f"AUDIT PIPELINE LOG\n")
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


# Глобальный логгер (инициализируется в main)
logger: PipelineLogger = None

# ============================================================================
# КОНСТАНТЫ И ПРОМПТЫ
# ============================================================================

# Модель по умолчанию
DEFAULT_MODEL = "gpt-4.1-mini"

# Количество параллельных запросов к LLM
MAX_WORKERS = 4

# Системный промпт для всех LLM-проверок
SYSTEM_PROMPT = """Ты — строгий аудитор документов.
Ты проверяешь ОДНО правило за один запрос.

Тебе даётся:
- фрагменты целевого документа (TARGET_*);
- фрагменты шаблона (TEMPLATE_*) — если требуется сверка;
- параметры правила: RULE_INDEX, COMPARE, SCOPE, RULE_TITLE;
- текстовые инструкции правила (RULE_INSTRUCTIONS).

COMPARE:
- template — сверяй чанк целевого с чанком шаблона (игнорируй плейсхолдеры: даты, ФИО, должности, названия компаний).
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
    """
    Спецификация одного правила проверки.
    """
    index: int
    title: str
    scope: Union[str, List[str]]  # Чанк(и) для проверки
    compare: str  # "target_only" | "template" | "cross_check"
    llm: bool  # True = LLM проверка, False = простая проверка
    instructions: List[str] = field(default_factory=list)


# ============================================================================
# ЗАГРУЗКА ПРАВИЛ
# ============================================================================

def load_rules(rules_path: str) -> List[RuleSpec]:
    """
    Загружает правила из JSON-файла.

    Args:
        rules_path: путь к файлу с правилами

    Returns:
        Список RuleSpec
    """
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
            instructions=rule.get("content", [])
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
    """
    Строит контекст для правила в зависимости от типа проверки.

    Args:
        spec: спецификация правила
        target_doc: распарсенный целевой документ
        template_doc: распарсенный шаблон

    Returns:
        Строка контекста для вставки в промпт
    """
    context_parts = []

    # Определяем scope (может быть строка или список)
    scopes = [spec.scope] if isinstance(spec.scope, str) else spec.scope

    if spec.compare == "target_only":
        # Только чанки целевого документа
        for scope in scopes:
            content = target_doc.get(scope, "")
            context_parts.append(f"[TARGET_{scope}]")
            context_parts.append(content)
            context_parts.append(f"[/TARGET_{scope}]")

    elif spec.compare == "template":
        # Чанки целевого + шаблона (попарно)
        for scope in scopes:
            # Целевой документ
            target_content = target_doc.get(scope, "")
            # Шаблон
            template_content = template_doc.get(scope, "")

            # Препроцессинг для правила #3: нормализуем плейсхолдеры
            if spec.index == 3 and scope == "текст_приказа":
                target_content = normalize_text_for_rule3(target_content)
                template_content = normalize_text_for_rule3(template_content)

            context_parts.append(f"[TARGET_{scope}]")
            context_parts.append(target_content)
            context_parts.append(f"[/TARGET_{scope}]")

            context_parts.append(f"[TEMPLATE_{scope}]")
            context_parts.append(template_content)
            context_parts.append(f"[/TEMPLATE_{scope}]")

    elif spec.compare == "cross_check":
        # Два или более чанков целевого документа для сравнения
        for scope in scopes:
            content = target_doc.get(scope, "")
            context_parts.append(f"[TARGET_{scope}]")
            context_parts.append(content)
            context_parts.append(f"[/TARGET_{scope}]")

    return "\n".join(context_parts)


def build_user_prompt(
    spec: RuleSpec,
    target_doc: Dict[str, Any],
    template_doc: Dict[str, Any]
) -> str:
    """
    Формирует пользовательский промпт для LLM.

    Args:
        spec: спецификация правила
        target_doc: распарсенный целевой документ
        template_doc: распарсенный шаблон

    Returns:
        Текст промпта
    """
    # Формируем строку scope
    scope_str = spec.scope if isinstance(spec.scope, str) else ", ".join(spec.scope)

    # Инструкции правила
    instructions_text = "\n".join(f"- {instr}" for instr in spec.instructions)

    # Контекст документов
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

def check_filename_rule(target_doc: Dict[str, Any]) -> List[Dict[str, Any]]:
    """
    Правило #1: проверка имени файла без LLM.

    Проверяет что имя файла содержит ожидаемый паттерн.
    """
    expected_pattern = "1.5 Приказ о создании ИЦ с прилож"
    actual = target_doc.get("имя_файла", "")
    actual_clean = actual.replace(".docx", "")

    # Проверяем вхождение паттерна
    if expected_pattern.lower() in actual_clean.lower():
        return []  # Норма

    return [{
        "rule_index": 1,
        "rule_title": "Проверка имени файла документа.",
        "Целевой документ": actual_clean,
        "Различие": f"Ожидалось имя файла, содержащее: «{expected_pattern}»"
    }]


def check_non_llm_rule(
    spec: RuleSpec,
    target_doc: Dict[str, Any]
) -> List[Dict[str, Any]]:
    """
    Выполняет проверку правила без LLM.

    Args:
        spec: спецификация правила
        target_doc: распарсенный целевой документ

    Returns:
        Список нарушений (пустой если норма)
    """
    # Правило #1 — проверка имени файла
    if spec.index == 1:
        return check_filename_rule(target_doc)

    # Если появятся другие non-LLM правила, добавить сюда

    return []


# ============================================================================
# LLM ПРОВЕРКИ
# ============================================================================

def parse_json_response(raw_response: str, spec_index: int, spec_title: str) -> List[Dict[str, Any]]:
    """
    Парсит JSON-ответ от LLM.

    Новый формат ответа — один объект с массивом "нарушения".
    Обрабатывает случаи когда LLM возвращает markdown код-блок.

    Args:
        raw_response: сырой ответ от LLM
        spec_index: номер правила (для заполнения в результате)
        spec_title: название правила (для заполнения в результате)

    Returns:
        Список нарушений (пустой если status=ok)
    """
    text = raw_response.strip()

    # Убираем markdown код-блок если есть
    if text.startswith("```"):
        lines = text.split("\n")
        if lines[0].startswith("```"):
            lines = lines[1:]
        if lines and lines[-1].strip() == "```":
            lines = lines[:-1]
        text = "\n".join(lines)

    # Санитизация: заменяем реальные переносы строк внутри JSON-строк на escape-последовательности
    # LLM иногда возвращает реальные \n внутри значений строк, что невалидно для JSON
    def sanitize_json_string(s: str) -> str:
        """Заменяет невалидные control characters внутри JSON-строк."""
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

            # Если мы внутри строки и встречаем control character — экранируем
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

        # Новый формат: объект с status
        if isinstance(result, dict):
            if result.get("status") == "ok":
                return []

            # status=fail — извлекаем нарушения
            violations = result.get("нарушения", [])
            # Добавляем rule_index и rule_title к каждому нарушению
            for v in violations:
                v["rule_index"] = result.get("rule_index", spec_index)
                v["rule_title"] = result.get("rule_title", spec_title)
            return violations

        # Старый формат (на случай если модель вернёт массив)
        if isinstance(result, list):
            return result

        return []

    except json.JSONDecodeError as e:
        print(f"[WARN] Не удалось распарсить JSON от LLM: {e}", file=sys.stderr)
        print(f"[WARN] Ответ: {text[:200]}...", file=sys.stderr)
        return []


def call_llm(
    messages: List[Dict[str, str]],
    model: str,
    temperature: float = 0.0
) -> str:
    """
    Вызывает LLM через OpenAI API.

    Args:
        messages: список сообщений
        model: название модели
        temperature: температура генерации

    Returns:
        Текст ответа
    """
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
    """
    Проверка одного правила через LLM.

    Args:
        spec: спецификация правила
        target_doc: распарсенный целевой документ
        template_doc: распарсенный шаблон
        model: модель для вызова
        temperature: температура генерации

    Returns:
        Список нарушений
    """
    global logger

    # Формируем сообщения
    user_prompt = build_user_prompt(spec, target_doc, template_doc)

    messages = [
        {"role": "system", "content": SYSTEM_PROMPT},
        {"role": "user", "content": user_prompt}
    ]

    # Логируем промпт
    if logger:
        logger.log_rule_prompt(spec.index, SYSTEM_PROMPT, user_prompt)

    # Вызываем LLM
    raw_response = call_llm(messages, model, temperature)

    # Парсим ответ (передаём spec для заполнения rule_index/title)
    violations = parse_json_response(raw_response, spec.index, spec.title)

    # Подстраховка — заполняем rule_index/title если не заполнено
    for obj in violations:
        obj.setdefault("rule_index", spec.index)
        obj.setdefault("rule_title", spec.title)

    # Логируем ответ
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
    temperature: float = 0.0,
    print_prompts: bool = False
) -> List[Dict[str, Any]]:
    """
    Выполняет все проверки: сначала non-LLM, затем LLM параллельно.

    Args:
        rules: список правил
        target_doc: распарсенный целевой документ
        template_doc: распарсенный шаблон
        model: модель для LLM
        temperature: температура генерации
        print_prompts: если True, только печатает промпты без вызова LLM

    Returns:
        Список всех нарушений
    """
    global logger
    all_violations = []

    # === Сначала правила без LLM (быстро, синхронно) ===
    for spec in rules:
        if not spec.llm:
            if logger:
                logger.log(f"🔧 Запуск non-LLM правила #{spec.index}: {spec.title}")
            violations = check_non_llm_rule(spec, target_doc)
            all_violations.extend(violations)

            # Логируем результат non-LLM проверки
            if logger:
                logger.log_non_llm_result(spec.index, violations)

            if print_prompts:
                print(f"\n{'='*60}")
                print(f"[NON-LLM] Правило #{spec.index}: {spec.title}")
                print(f"Результат: {violations}")

    # === Затем правила с LLM ===
    llm_rules = [r for r in rules if r.llm]

    if print_prompts:
        # Режим отладки — только печать промптов
        for spec in llm_rules:
            user_prompt = build_user_prompt(spec, target_doc, template_doc)
            print(f"\n{'='*60}")
            print(f"[LLM] Правило #{spec.index}: {spec.title}")
            print(f"{'='*60}")
            print("\n--- SYSTEM PROMPT ---")
            print(SYSTEM_PROMPT)
            print("\n--- USER PROMPT ---")
            print(user_prompt)

            # Также сохраняем в лог если логгер есть
            if logger:
                logger.log_rule_prompt(spec.index, SYSTEM_PROMPT, user_prompt)
        return all_violations

    # Параллельное выполнение LLM-проверок
    if logger:
        logger.log(f"🚀 Запуск {len(llm_rules)} LLM-проверок параллельно (max_workers={MAX_WORKERS})")

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

    # Сортируем по номеру правила
    return sorted(all_violations, key=lambda x: x.get("rule_index", 0))


# ============================================================================
# СОХРАНЕНИЕ РЕЗУЛЬТАТОВ
# ============================================================================

def save_to_excel(
    violations: List[Dict[str, Any]],
    output_path: str
) -> None:
    """
    Сохраняет результаты в Excel файл.

    Args:
        violations: список нарушений
        output_path: путь для сохранения
    """
    if not violations:
        # Создаём пустой файл с заголовками
        df = pd.DataFrame(columns=[
            "rule_index",
            "rule_title",
            "Целевой документ",
            "Различие"
        ])
    else:
        df = pd.DataFrame(violations)
        # Упорядочиваем колонки
        columns_order = ["rule_index", "rule_title", "Целевой документ", "Различие"]
        existing_cols = [c for c in columns_order if c in df.columns]
        extra_cols = [c for c in df.columns if c not in columns_order]
        df = df[existing_cols + extra_cols]

    # Создаём директорию если нужно
    Path(output_path).parent.mkdir(parents=True, exist_ok=True)

    # Сохраняем
    df.to_excel(output_path, index=False, engine='openpyxl')
    print(f"✅ Результат сохранён: {output_path}", file=sys.stderr)


# ============================================================================
# MAIN
# ============================================================================

def main():
    """Точка входа CLI."""
    global logger

    parser = argparse.ArgumentParser(
        description="LLM-аудит документа 'Приказ о создании ИЦ'"
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
        help="Путь к JSON с правилами (по умолчанию: config/tz_prikaz_ic.json)"
    )
    parser.add_argument(
        "--model",
        default=DEFAULT_MODEL,
        help=f"Модель OpenAI (по умолчанию: {DEFAULT_MODEL})"
    )
    parser.add_argument(
        "--out-xlsx",
        default=None,
        help="Путь для сохранения результата в Excel"
    )
    parser.add_argument(
        "--print-prompts",
        action="store_true",
        help="Отладка: вывести промпты без вызова LLM"
    )
    parser.add_argument(
        "--temperature",
        type=float,
        default=0.0,
        help="Температура генерации LLM (по умолчанию: 0.0)"
    )
    parser.add_argument(
        "--session-dir",
        default=None,
        help="Директория сессии для всех артефактов (по умолчанию: logs_result/session_YYYYMMDD_HHMMSS)"
    )
    parser.add_argument(
        "--target-json",
        default=None,
        help="Путь к готовому JSON целевого документа (для тестирования, вместо парсинга .docx)"
    )

    args = parser.parse_args()

    # Определяем путь к правилам
    if args.rules:
        rules_path = args.rules
    else:
        script_dir = Path(__file__).parent.parent
        rules_path = script_dir / "config" / "tz_prikaz_ic.json"

    # Инициализируем директорию сессии (все артефакты в одном месте)
    if args.session_dir:
        session_dir = Path(args.session_dir)
    else:
        script_dir = Path(__file__).parent.parent
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        session_dir = script_dir / "logs_result" / f"session_{timestamp}"

    logger = PipelineLogger(session_dir)
    logger.log(f"🚀 Запуск аудита")
    logger.log(f"   Сессия: {session_dir}")
    logger.log(f"   Целевой документ: {args.target}")
    logger.log(f"   Шаблон: {args.template}")
    logger.log(f"   Правила: {rules_path}")
    logger.log(f"   Модель: {args.model}")
    logger.log(f"   Температура: {args.temperature}")

    # Проверяем существование файлов
    if not args.target_json and not Path(args.target).exists():
        logger.log(f"❌ Целевой документ не найден: {args.target}")
        sys.exit(1)

    if not Path(args.template).exists():
        logger.log(f"❌ Шаблон не найден: {args.template}")
        sys.exit(1)

    if not Path(rules_path).exists():
        logger.log(f"❌ Файл правил не найден: {rules_path}")
        sys.exit(1)

    # Парсим документы (или загружаем готовый JSON для тестов)
    if args.target_json:
        logger.log(f"📄 Загрузка целевого документа из JSON (тестовый режим): {args.target_json}")
        with open(args.target_json, 'r', encoding='utf-8') as f:
            target_doc = json.load(f)
        logger.log_parsed_doc(target_doc, "target_doc_from_json")
    else:
        logger.log(f"📄 Парсинг целевого документа...")
        target_doc = parse_prikaz_ic(args.target)
        logger.log_parsed_doc(target_doc, "target_doc")

    logger.log(f"📄 Парсинг шаблона...")
    template_doc = parse_prikaz_ic(args.template)
    logger.log_parsed_doc(template_doc, "template_doc")

    # Загружаем правила
    logger.log(f"📋 Загрузка правил...")
    rules = load_rules(str(rules_path))
    logger.log(f"   Загружено {len(rules)} правил")

    # Сохраняем правила в лог
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
        temperature=args.temperature,
        print_prompts=args.print_prompts
    )

    # Логируем финальные результаты
    logger.log_final_results(violations)

    # Выводим JSON результат в stdout
    print(json.dumps(violations, ensure_ascii=False, indent=2))

    # Сохраняем Excel в директорию сессии
    if args.out_xlsx:
        xlsx_path = args.out_xlsx
    else:
        # По умолчанию сохраняем в директорию сессии
        xlsx_path = str(session_dir / "audit_result.xlsx")
    save_to_excel(violations, xlsx_path)
    logger.log(f"📊 Excel сохранён: {xlsx_path}")

    # Итоговая статистика
    logger.log(f"{'='*60}")
    logger.log(f"📊 Итого нарушений: {len(violations)}")
    if violations:
        # Группируем по правилам
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
