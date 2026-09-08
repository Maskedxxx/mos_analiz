# START_MODULE_CONTRACT
# PURPOSE: Multi-rule LLM-аудит — единственный активный путь отправки документа в LLM. Документ целиком + семантическая карта секций + список правил → один запрос к Qwen → массив вердиктов → список нарушений.
# INPUTS: `parsed` (ParsedDocument с raw_text), `sections` (dict от `sections.json`), `rules` (list от `rules_multi.json` / `rules_methodology.json`), `filename`, `llm_base_url`, `llm_model`, опциональная `session_dir` для debug-артефактов.
# OUTPUTS: `{verdicts, violations, unchecked, checked_count, usage, prompt_chars}`.
# KEYWORDS: multi-rule, llm, qwen, audit, system-prompt, user-prompt-template.
# LINKS: config/llm.py (LLM_CONFIG), src/llm/client.py (call_llm — не используется тут, т.к. multi_rule имеет Qwen-специфичные параметры), main.py::AuditEngine.run, doc_configs/<тип>/rules_multi.json, doc_configs/<тип>/rules_methodology.json.
# RATIONALE: Единственный активный LLM-путь. Остальные (single-rule, template comparison, scope-based rules) — мёртвое legacy после step 1. SYSTEM_PROMPT и шаблоны user-промпта вынесены как module-level константы: глядя на верх файла, сразу видно, что именно уходит в Qwen.
# END_MODULE_CONTRACT

from __future__ import annotations

# START_IMPORTS
import json
import logging
import re
from pathlib import Path
from typing import Any, Dict, List, Optional


from config.llm import LLM_CONFIG
from src.llm.client import make_llm_client
# END_IMPORTS


# START_LOGGER
logger = logging.getLogger(__name__)
# END_LOGGER


# START_CONTEXT_GUARD
# PURPOSE: Оценка размера промпта до вызова модели (F10, находка 2.4g). Точного токенизатора на
# стороне сервиса нет; берём консервативную оценку «1 токен ≈ 3 символа» (для кириллицы Qwen
# даёт больше символов на токен, так что оценка завышена — лучше отклонить редкий гигантский
# документ заранее, чем получить сырой 400 от llama.cpp). Лимит — LLM_CONFIG.context_tokens.
_CHARS_PER_TOKEN = 3
_CHARS_PER_PAGE = 3000  # для подсказки «≈N страниц» в тексте ошибки


def check_prompt_fits_context(prompt_chars: int, context_tokens: int) -> None:
    """Бросает ValueError с человекочитаемым текстом, если промпт не помещается в контекст модели."""
    est_tokens = prompt_chars // _CHARS_PER_TOKEN
    if est_tokens > context_tokens:
        pages = max(1, prompt_chars // _CHARS_PER_PAGE)
        max_pages = max(1, (context_tokens * _CHARS_PER_TOKEN) // _CHARS_PER_PAGE)
        raise ValueError(
            f"Документ слишком большой для проверки: примерно {pages} страниц текста, "
            f"максимум около {max_pages}. Разделите документ или уберите приложения."
        )
# END_CONTEXT_GUARD


# START_SYSTEM_PROMPT
# PURPOSE: Системный промпт для multi-rule аудита. Инструктирует LLM формат ответа (строго JSON-массив, по одному вердикту на правило, reasoning + verdict в каждом).
# INPUTS: —
# OUTPUTS: Константа, используемая как `role: system` в `run_multi_rule_audit`.
# KEYWORDS: system-prompt, qwen, json-output, verdict, reasoning.
SYSTEM_PROMPT = """\
Ты — строгий аудитор документов.

═══════════════════════════════════════════════════════════════════════
КОНТРАКТ ОТВЕТА (ЖЁСТКИЙ, НАРУШЕНИЕ = ОШИБКА АУДИТА):
- На входе ты получишь список из N правил (количество указано в заголовке `## RULES (N=…)`).
- На выходе ОБЯЗАН вернуть JSON-массив из РОВНО N объектов — по одному на каждое правило, в том же порядке.
- Каждое правило ОБЯЗАТЕЛЬНО получает свой объект, даже если документ короткий, повторяется или ты считаешь, что нечего добавить:
    • Если требование выполнено — status="ok" и краткий reasoning.
    • Если требование не выполнено — status="fail" + нарушения.
    • Если секции физически нет — status="error" + reason.
- ЗАПРЕЩЕНО: возвращать один объект вместо массива. Останавливаться после первого правила. Пропускать правила «потому что нечего сказать». Объединять несколько правил в один объект.
- ПЕРЕД ОТПРАВКОЙ ОТВЕТА посчитай количество объектов в массиве — оно должно равняться N из заголовка `## RULES (N=…)`. Если не равно — переделай ответ.
═══════════════════════════════════════════════════════════════════════

В одном запросе тебе даётся ВЕСЬ документ и СПИСОК ПРАВИЛ для проверки.

Документ логически разделён на именованные СЕКЦИИ. Для каждой секции даны границы (маркеры начала/конца) и описание. Секции НЕ размечены в тексте — ты должен определить их сам по маркерам.

Для каждого правила указана ЦЕЛЕВАЯ СЕКЦИЯ (target_section). Анализируй ТОЛЬКО эту секцию документа, даже если похожий текст встречается в других местах. Если у правила несколько target_sections — проверяй их совокупно (достаточно выполнения в любой из них, если явно не сказано иное).

ПОРЯДОК РАБОТЫ ПО КАЖДОМУ ПРАВИЛУ:
1. Найди в документе target_section правила.
2. Проверь требования из «check» (с учётом «exclusions», если они есть).
3. Сначала запиши «reasoning» — обоснование: что увидел в секции, какие требования выполнены, какие нет.
4. Затем поставь «status» — он ОБЯЗАН прямо следовать из reasoning, не противоречить ему.

СЕМАНТИКА СТАТУСОВ (запомни строго):
- «ok» — все требования из «check» выполнены в указанной секции. Если в reasoning ты констатировал «всё на месте / соответствует / заполнено», статус — «ok».
- «fail» — хотя бы одно требование из «check» НЕ выполнено: отсутствует, нарушено, заменено плейсхолдером (подчёркивания, поясняющий текст вместо значения, пустые кавычки и т.п.). Если в reasoning ты констатировал «отсутствует / плейсхолдер / не заполнено / не соответствует / не найдено», статус — «fail», а не «ok».
- «error» — секцию физически невозможно проверить (нужная секция отсутствует в документе вовсе). Добавь поле «reason».

ПРОВЕРКА КОНСИСТЕНТНОСТИ перед записью verdict: перечитай свой reasoning и убедись, что status логически из него следует. Противоречие («reasoning описывает нарушение, но status=ok») — недопустимо.

Формат ответа — СТРОГО JSON-массив, ПО ОДНОМУ объекту на правило, в том же порядке, что и в списке правил. У каждого объекта ДВА обязательных поля: «reasoning» и «verdict».

ДЕТАЛИЗАЦИЯ НАРУШЕНИЙ (важно для отчёта):
- Поле «нарушения» — массив. Если в правиле check состоит из НЕСКОЛЬКИХ требований (например, «(1) ..., (2) ..., (3) ...») и часть из них не выполнена — выписывай КАЖДОЕ невыполненное требование отдельным объектом в массиве «нарушения», а не сворачивай их в одну строку.
- Запрещено сворачивать множественные пропуски в одну формулировку типа «не описаны мотивация, комиссия, оценка...». Правильно: 3 отдельных объекта.
- Каждый объект «нарушения» проверяет одну атомарную мысль из check.

Пример ответа для списка из 2 правил (демонстрирует и оборачивание в массив, и детализацию нарушений):
[
  {
    "rule_index": 1,
    "reasoning": "В секции X выполнены все три требования: (1) ..., (2) ..., (3) ...",
    "verdict": {
      "status": "ok"
    }
  },
  {
    "rule_index": 3,
    "reasoning": "В секции Y из 4 подпунктов check выполнен только (1), пропущены (2), (3), (4).",
    "verdict": {
      "status": "fail",
      "нарушения": [
        {"Целевой документ": "<цитата или 'отсутствует'>", "Различие": "не выполнен подпункт (2) — <что именно>"},
        {"Целевой документ": "<цитата или 'отсутствует'>", "Различие": "не выполнен подпункт (3) — <что именно>"},
        {"Целевой документ": "<цитата или 'отсутствует'>", "Различие": "не выполнен подпункт (4) — <что именно>"}
      ]
    }
  }
]

Поля:
- «rule_index» — номер правила из списка (обязательно)
- «reasoning» — КРАТКОЕ рассуждение (1-3 предложения): какую секцию смотрел, что нашёл/не нашёл, почему такой вердикт. Без пересказа содержимого документа и правила.
- «verdict» — объект с финальным вердиктом (ТОЛЬКО структурированный, никаких рассуждений):
    - «status»: «ok» / «fail» / «error»
    - «нарушения»: массив {«Целевой документ», «Различие»}, только при «fail»
    - «Целевой документ»: цитата из документа или «отсутствует»
    - «Различие»: что не так или чего не хватает
- «error» — status для случая, когда правило технически невозможно проверить (нет нужной секции в документе). Добавь поле «reason».

ВАЖНО:
- Ответ — ТОЛЬКО JSON-массив, без markdown-ограждения и без пояснений до/после массива.
- В поле «verdict» — СТРОГО структурированные данные. Никаких «проверю ещё раз», «пересмотрю». Все такие мысли — в «reasoning».
- В «reasoning» НЕ включай JSON-объекты и не пытайся там формировать ответ — это свободный текст для размышлений.
- Не пропускай правила — в массиве должно быть ровно столько объектов, сколько правил.
- Не смешивай правила между собой.
- Игнорируй OCR-артефакты (пробелы между буквами, склейку строк, дублирование) — оценивай смысл.
- Уважай target_section правила: не ищи нарушения вне указанной секции.
- ВАЛИДНЫЙ JSON: внутри строковых значений ВСЕ двойные кавычки экранируй обратной чертой. Пример: «ООО "ВЭЙВ"» внутри reasoning записывай либо как «ООО \\"ВЭЙВ\\"», либо лучше используй ёлочки «ООО «ВЭЙВ»» / одинарные кавычки 'ВЭЙВ'. Неэкранированные двойные кавычки ломают парсер и твой ответ теряется.
- ИСТОЧНИК ПРАВИЛА: если у правила указаны source_ref и/или nature — обязательно упомяни их в КОНЦЕ reasoning в формате `[Источник: <source_ref>, носит <nature> характер]`. Это нужно чтобы эксперт мог быстро найти первоисточник в методичке (для юр.отдела). При конфликте требований МУ vs МР — МУ имеет приоритет (МУ = «обязательно», МР = «рекомендательно»)."""
# END_SYSTEM_PROMPT


# START_USER_PROMPT_TEMPLATES
# PURPOSE: Шаблоны user-промпта и его блоков (SECTIONS/RULES). Именно они через `.format(...)` собирают текст, уходящий в LLM. Single source of truth.
# INPUTS: placeholders.
# OUTPUTS: строковые шаблоны.
# KEYWORDS: user-prompt, template, sections, rules.

# ───────────────────────────────────────────────────────────────────────────────
# Итоговый user-промпт multi-rule запроса:
#
#   ## FILENAME
#   <имя файла>
#
#   ## DOCUMENT
#   <raw_text всего документа>
#
#   ## SECTIONS
#   - **<имя_секции>**: <описание>
#       start: <маркер начала>
#       end: <маркер конца>
#   ... (по одной записи на секцию)
#
#   ## RULES
#   ### RULE <index> — <title>
#   target_section: <имя секции>      ← ИЛИ target_sections: sec1, sec2
#   check: <что проверять>
#   exclusions: <когда правило не применяется>  ← опционально
#   ... (по одному блоку на правило)
# ───────────────────────────────────────────────────────────────────────────────
USER_PROMPT_TEMPLATE = """\
## FILENAME
{filename}

## DOCUMENT
{doc_text}

## SECTIONS
{sections_block}

## RULES (N={rules_count})
ВНИМАНИЕ: ниже {rules_count} правил. Ответ ДОЛЖЕН быть JSON-массивом из РОВНО {rules_count} объектов, по одному на каждое правило, в том же порядке.

{rules_block}
"""

SECTION_BLOCK_TEMPLATE = """\
- **{name}**: {description}
    start: {start}
    end: {end}"""

# Один блок RULE. Поля `target_line`, `source_ref_line`, `nature_line` и
# `exclusions_line` могут быть пустыми — пустая строка не добавляет лишнюю пустую
# строку в финальный текст, т.к. они уже собираются через \n-join из непустых
# элементов. source_ref/nature — источник правила (МР/МУ) для отчёта эксперту.
RULE_BLOCK_TEMPLATE = """\
### RULE {index} — {title}
{target_line}{source_ref_line}{nature_line}
check: {check}{exclusions_line}
"""
# END_USER_PROMPT_TEMPLATES


# START_BUILD_USER_PROMPT
# PURPOSE: Собирает финальный user-промпт по шаблонам выше.
# INPUTS: `filename`, `doc_text` (raw_text документа), `sections` (dict от sections.json), `rules` (list правил).
# OUTPUTS: Готовый текст user-промпта.
# KEYWORDS: build-user-prompt, format-template.
def build_user_prompt(filename: str, doc_text: str, sections: Dict[str, Any], rules: List[Dict[str, Any]]) -> str:
    """
    Назначение:
        Собирает user-промпт из шаблонов `USER_PROMPT_TEMPLATE` / `SECTION_BLOCK_TEMPLATE` /
        `RULE_BLOCK_TEMPLATE`.

    Вход:
        filename: Имя файла (попадает в ## FILENAME).
        doc_text: Весь текст документа (попадает в ## DOCUMENT).
        sections: `{имя → {description, start, end}}` — семантическая карта секций.
        rules: Список правил. Каждое `{index, title, check, target_section|target_sections, exclusions?}`.

    Выход:
        Готовый текст промпта.

    Логика:
        1. Форматирует каждую секцию через `SECTION_BLOCK_TEMPLATE`.
        2. Для каждого правила определяет строку target:
           `target_section: <name>` ИЛИ `target_sections: name1, name2`.
           Если правило не задало ни одно — строка пустая.
        3. Форматирует правило через `RULE_BLOCK_TEMPLATE`, добавляет exclusions только
           если оно есть в правиле (иначе пустая строка).
        4. Склеивает всё через `USER_PROMPT_TEMPLATE`.
    """
    sections_block = "\n".join(
        SECTION_BLOCK_TEMPLATE.format(
            name=name,
            description=meta["description"],
            start=meta["start"],
            end=meta["end"],
        )
        for name, meta in sections.items()
    )

    rule_blocks: List[str] = []
    for rule in rules:
        if "target_section" in rule:
            target_line = f"target_section: {rule['target_section']}"
        elif "target_sections" in rule:
            target_line = f"target_sections: {', '.join(rule['target_sections'])}"
        else:
            target_line = ""
        # Источник правила (МР/МУ + пункт) и его природа (обязательно/рекомендательно) —
        # для показа в reasoning и отчёте эксперту. Только если заданы в правиле.
        source_ref_line = f"\nsource_ref: {rule['source_ref']}" if rule.get("source_ref") else ""
        nature_line = f"\nnature: {rule['nature']}" if rule.get("nature") else ""
        exclusions_line = f"\nexclusions: {rule['exclusions']}" if rule.get("exclusions") else ""
        rule_blocks.append(RULE_BLOCK_TEMPLATE.format(
            index=rule["index"],
            title=rule["title"],
            target_line=target_line,
            source_ref_line=source_ref_line,
            nature_line=nature_line,
            check=rule["check"],
            exclusions_line=exclusions_line,
        ))
    rules_block = "\n".join(rule_blocks)

    return USER_PROMPT_TEMPLATE.format(
        filename=filename,
        doc_text=doc_text,
        sections_block=sections_block,
        rules_count=len(rules),
        rules_block=rules_block,
    )
# END_BUILD_USER_PROMPT


# START_COLLECT_DOC_TEXT
# PURPOSE: Достаёт текст документа из `ParsedDocument` для подстановки в ## DOCUMENT.
# INPUTS: `parsed` (dict от парсера), `include_scopes` (опционально, легаси).
# OUTPUTS: Текст документа.
# KEYWORDS: raw-text, legacy-fallback.
def _collect_doc_text(parsed: Dict[str, Any], include_scopes: Optional[List[str]]) -> str:
    """
    Назначение:
        Возвращает текст документа для передачи в LLM.

    Вход:
        parsed: Результат парсера. Активный путь — `ParsedDocument{filename, path, raw_text}`.
        include_scopes: Опциональный список scope-имён для фильтрации (легаси).

    Выход:
        Текст документа.

    Логика:
        1. **Активный путь:** если `parsed['raw_text']` — непустая строка, возвращаем
           её целиком. Это основной сценарий для docx/pptx/pdf парсеров.
        2. **Legacy-путь:** если raw_text нет (старые парсеры со scope-словарями) —
           собираем непустые scope-значения с дедупликацией. Оставлено для
           совместимости, пока не будут удалены все legacy-парсеры (сейчас уже нет
           таких в активном рантайме, но код безвреден).
    """
    raw_text = parsed.get("raw_text")
    if isinstance(raw_text, str) and raw_text:
        return raw_text
    if include_scopes:
        text_keys = [k for k in include_scopes if k in parsed]
    else:
        text_keys = [k for k in parsed.keys() if k not in ("filename", "path")]
    unique_texts = list(dict.fromkeys(parsed[k] for k in text_keys if isinstance(parsed.get(k), str)))
    return "\n\n".join(unique_texts) if unique_texts else ""
# END_COLLECT_DOC_TEXT


# START_PARSE_RESPONSE
# PURPOSE: Парсинг JSON-массива вердиктов от LLM. Несколько fallback-уровней, потому что LLM иногда ломает структуру между объектами.
# INPUTS: Сырой текст ответа.
# OUTPUTS: Список вердиктов (dict'ов).
# KEYWORDS: json-parse, markdown-strip, object-splitter, resilience.
def _parse_response(response_text: str) -> List[Dict[str, Any]]:
    """
    Назначение:
        Достаёт список вердиктов из ответа LLM. Обрабатывает markdown-обёртку и
        пытается восстановить валидные объекты, даже если LLM склеил/сломал структуру.

    Вход:
        response_text: Сырой ответ LLM.

    Выход:
        Список вердиктов — каждый `{rule_index, reasoning, verdict}`.

    Логика:
        1. Снимает ```json ... ``` обёртку если есть.
        2. Пытается `json.loads` — если list, возвращает; если dict, возвращает [dict].
        3. Fallback 1: сканирует посимвольно, выделяет объекты верхнего уровня по
           балансу скобок и парсит каждый отдельно.
        4. Fallback 2: ищет позиции `"rule_index"`, восстанавливает границы объектов
           и компенсирует дисбаланс скобок (`{` vs `}`).
    """
    clean = response_text.strip()
    if clean.startswith("```"):
        parts = clean.split("```", 2)
        if len(parts) >= 2:
            clean = parts[1]
            if clean.lstrip().startswith("json"):
                clean = clean.lstrip()[4:]
    clean = clean.strip()
    try:
        parsed = json.loads(clean)
        if isinstance(parsed, list):
            return parsed
        if isinstance(parsed, dict):
            return [parsed]
    except json.JSONDecodeError:
        pass

    # Fallback 1: сканирование посимвольно.
    verdicts: List[Dict[str, Any]] = []
    i = 0
    while i < len(clean):
        if clean[i] != "{":
            i += 1
            continue
        depth = 0
        start = i
        in_str = False
        escape = False
        while i < len(clean):
            ch = clean[i]
            if escape:
                escape = False
            elif ch == "\\":
                escape = True
            elif ch == '"':
                in_str = not in_str
            elif not in_str:
                if ch == "{":
                    depth += 1
                elif ch == "}":
                    depth -= 1
                    if depth == 0:
                        i += 1
                        break
            i += 1
        obj_text = clean[start:i]
        try:
            verdicts.append(json.loads(obj_text))
        except json.JSONDecodeError:
            logger.warning(f"Пропущен невалидный JSON-объект: {obj_text[:100]}...")
    if verdicts:
        return verdicts

    # Fallback 2: извлечение по маркерам `"rule_index"`.
    rule_markers = [m.start() for m in re.finditer(r'"rule_index"\s*:', clean)]
    if len(rule_markers) >= 2:
        boundaries: List[int] = []
        for marker_pos in rule_markers:
            brace = clean.rfind("{", 0, marker_pos)
            if brace >= 0:
                boundaries.append(brace)
        boundaries.append(len(clean))
        for idx in range(len(boundaries) - 1):
            start = boundaries[idx]
            end = boundaries[idx + 1]
            chunk = clean[start:end]
            last_brace = chunk.rfind("}")
            if last_brace < 0:
                continue
            candidate = chunk[: last_brace + 1].strip().rstrip(",").strip()
            opens = candidate.count("{")
            closes = candidate.count("}")
            if closes > opens:
                extra = closes - opens
                for _ in range(extra):
                    candidate = candidate.rstrip()
                    if candidate.endswith("}"):
                        candidate = candidate[:-1].rstrip()
            elif closes < opens:
                candidate += "}" * (opens - closes)
            try:
                verdicts.append(json.loads(candidate))
            except json.JSONDecodeError:
                logger.warning(f"Пропущен чанк: {candidate[:80]}...")
    if not verdicts:
        raise json.JSONDecodeError("Не удалось распарсить ни одного объекта", clean, 0)
    return verdicts
# END_PARSE_RESPONSE


# START_VERDICT_TO_VIOLATIONS
# PURPOSE: Конвертирует LLM-вердикты в legacy-формат violations, который ожидает UI/Excel-репортёр.
# INPUTS: Список вердиктов, список правил (для rule_title), `layer` (base / methodology).
# OUTPUTS: Список нарушений.
# KEYWORDS: verdict-to-violation, legacy-format, layer.
def _verdict_to_violations(verdicts: List[Dict[str, Any]], rules: List[Dict[str, Any]], layer: str = "base") -> List[Dict[str, Any]]:
    """
    Назначение:
        Приводит ответ multi-rule к legacy-формату нарушений.

    Вход:
        verdicts: Ответ LLM (список dict'ов).
        rules: Правила (для подстановки `title` по `rule_index`).
        layer: `base` (менеджерский слой) или `methodology` (из МР/МУ).

    Выход:
        Список нарушений `[{правило, rule_index, layer, Целевой документ, Различие}]`.

    Логика:
        Только вердикты со `status == fail` превращаются в нарушения; внутри каждого
        `нарушения` — массив конкретных инстансов, каждый добавляется как отдельная
        запись.
    """
    rules_by_idx = {r["index"]: r for r in rules}
    violations: List[Dict[str, Any]] = []
    for v in verdicts:
        idx = v.get("rule_index")
        rule = rules_by_idx.get(idx, {})
        rule_title = rule.get("title", f"Правило {idx}")
        # Источник правила (МР/МУ + пункт) и его природа — для отчёта эксперту (юр.отдел).
        source_ref = rule.get("source_ref", "")
        nature = rule.get("nature", "")
        verdict_obj = v.get("verdict", {}) if isinstance(v.get("verdict"), dict) else {}
        status = verdict_obj.get("status", "?")
        # reasoning модели — для колонки «Обоснование» в Excel; хвост «[Источник: …]» срезаем,
        # т.к. источник уже выводится в отдельной колонке «Источник (МР/МУ)».
        reasoning = v.get("reasoning", "") or ""
        if isinstance(reasoning, str):
            _pos = reasoning.rfind("[Источник:")
            if _pos != -1:
                reasoning = reasoning[:_pos].rstrip()
        if status == "fail":
            for violation in verdict_obj.get("нарушения", []):
                violations.append({
                    "правило": rule_title,
                    "rule_index": idx,
                    "layer": layer,
                    "Источник": source_ref,
                    "Природа": nature,
                    "Целевой документ": violation.get("Целевой документ", "отсутствует"),
                    "Различие": violation.get("Различие", "?"),
                    "Обоснование": reasoning,
                })
    return violations
# END_VERDICT_TO_VIOLATIONS


# START_RUN_MULTI_RULE_AUDIT
# PURPOSE: Главная функция — оркестрирует один multi-rule прогон: build_user_prompt → LLM → parse_response → verdict_to_violations.
# INPUTS: parsed, sections, rules, include_scopes, filename, llm_base_url, llm_model, session_dir, layer.
# OUTPUTS: `{verdicts, violations, usage, prompt_chars}`.
# KEYWORDS: run-audit, qwen-call, session-artifacts.
def run_multi_rule_audit(
    *,
    parsed: Dict[str, Any],
    sections: Dict[str, Dict[str, str]],
    rules: List[Dict[str, Any]],
    include_scopes: Optional[List[str]],
    filename: str,
    llm_base_url: str,
    llm_model: Optional[str] = None,
    session_dir: Optional[Path] = None,
    layer: str = "base",
) -> Dict[str, Any]:
    """
    Назначение:
        Запускает multi-rule audit: сборка промпта → LLM → парсинг вердикта →
        формирование нарушений.

    Вход:
        parsed: Результат парсера (`ParsedDocument`).
        sections: Карта секций (`sections.json`).
        rules: Список правил (`rules_multi.json` или `rules_methodology.json`).
        include_scopes: Легаси-фильтр (сейчас не используется при наличии raw_text).
        filename: Имя файла для подстановки в промпт.
        llm_base_url: URL LLM сервиса (обычно из `AuditConfig.llm_base_url` или
            `LLM_CONFIG.base_url`).
        llm_model: Имя модели. `None` → `LLM_CONFIG.default_model` (env LLM_MODEL).
        session_dir: Если задан — сохраняет debug-артефакты
            (system_prompt.txt, user_prompt.txt, response_raw.txt, response_parsed.json).
        layer: `base` или `methodology` — префикс для имён debug-файлов и поле
            нарушений.

    Выход:
        Dict `{verdicts, violations, unchecked, checked_count, usage, prompt_chars}`.
        `unchecked` — правила без годного вердикта после всех попыток (F15).

    Логика:
        1. Достаёт текст документа через `_collect_doc_text`.
        2. Собирает user-промпт через `build_user_prompt`.
        3. Если `session_dir` задан — сохраняет system_prompt + user_prompt.
        4. Делает прямой chat.completions.create с Qwen-специфичными параметрами
           (top_p, top_k, min_p, repetition_penalty, chat_template_kwargs.enable_thinking=False).
           Эти значения **калиброваны для Qwen multi-rule аудита** и не должны
           меняться без тестирования качества.
        5. Сохраняет raw response если session_dir.
        6. Парсит через `_parse_response`, сохраняет parsed JSON.
        7. Конвертирует в violations через `_verdict_to_violations`.
    """
    doc_text = _collect_doc_text(parsed, include_scopes)
    user_prompt = build_user_prompt(filename, doc_text, sections, rules)
    prefix = "multi_rule" if layer == "base" else f"multi_rule_{layer}"
    if session_dir:
        session_dir = Path(session_dir)
        session_dir.mkdir(parents=True, exist_ok=True)
        (session_dir / f"{prefix}_system_prompt.txt").write_text(SYSTEM_PROMPT, encoding="utf-8")
        (session_dir / f"{prefix}_user_prompt.txt").write_text(user_prompt, encoding="utf-8")

    # F10: документ больше контекста — понятный отказ до обращения к модели.
    check_prompt_fits_context(len(SYSTEM_PROMPT) + len(user_prompt), LLM_CONFIG.context_tokens)

    llm_model = llm_model or LLM_CONFIG.default_model
    client = make_llm_client(llm_base_url)
    # Ретрай при «каше» ответа LLM: модель иногда возвращает не JSON-массив вердиктов,
    # а одиночный объект/ошибку → парсер даёт < N годных вердиктов и весь слой молча обнуляется.
    # Переспрашиваем модель, меняя seed (иначе повтор даст тот же результат), берём лучшую попытку.
    max_attempts = 5  # число попыток при неполном ответе (крупные доки чаще обрезаются)

    def _valid_count(vs: List[Dict[str, Any]]) -> int:
        # Годный вердикт = dict с распознаваемым rule_index (int/строка-цифра)
        # И непустым verdict.status (ok/fail/error). Пустой/битый verdict (напр. null) —
        # тоже «каша» ответа и должен триггерить ретрай, а не считаться годным.
        n = 0
        for v in vs:
            if not isinstance(v, dict):
                continue
            ri = v.get("rule_index")
            has_ri = isinstance(ri, int) or (isinstance(ri, str) and ri.strip().isdigit())
            verd = v.get("verdict")
            status = verd.get("status") if isinstance(verd, dict) else None
            has_status = isinstance(status, str) and status.strip() != ""
            if has_ri and has_status:
                n += 1
        return n

    expected = len(rules)
    best_verdicts: List[Dict[str, Any]] = []
    best_text = ""
    best_response = None
    response = None
    for attempt in range(max_attempts):
        response = client.chat.completions.create(
            model=llm_model,
            messages=[
                {"role": "system", "content": SYSTEM_PROMPT},
                {"role": "user", "content": user_prompt},
            ],
            temperature=0.7,
            top_p=0.8,
            max_tokens=64000,
            seed=42 + attempt,
            extra_body={
                "chat_template_kwargs": {"enable_thinking": False},
                "top_k": 20,
                "min_p": 0.0,
                "repetition_penalty": 1.0,
            },
        )
        response_text = response.choices[0].message.content or ""
        # F16: нечитаемый ответ (текст отказа, пустой content, обрезанный JSON) — та же «каша»,
        # что и неполный массив: не падаем на первой попытке, а идём на следующий seed.
        try:
            verdicts = _parse_response(response_text)
        except json.JSONDecodeError as e:
            logger.warning(f"multi_rule[{layer}]: попытка {attempt + 1}/{max_attempts} — ответ не разобран ({e}) — ретрай со сменой seed")
            verdicts = []
        if _valid_count(verdicts) > _valid_count(best_verdicts):
            best_verdicts = verdicts
            best_text = response_text
            best_response = response
        if _valid_count(verdicts) >= expected:
            break
        logger.warning(
            f"multi_rule[{layer}]: попытка {attempt + 1}/{max_attempts} дала "
            f"{_valid_count(verdicts)}/{expected} годных вердиктов (каша ответа) — ретрай со сменой seed"
        )

    verdicts = best_verdicts
    response_text = best_text
    if best_response is not None:
        response = best_response
    elif _valid_count(best_verdicts) == 0:
        # Ни одна из попыток не дала разбираемого ответа — говорим по-человечески, без JSONDecodeError.
        # Сырой ответ последней попытки сохраняем для разбора администратором.
        if session_dir:
            (session_dir / f"{prefix}_response_raw.txt").write_text(response.choices[0].message.content or "", encoding="utf-8")
        raise ValueError(
            f"Модель не вернула результат в ожидаемом формате после {max_attempts} попыток. "
            "Повторите проверку позже; если повторяется — обратитесь к администратору."
        )

    # F15: какие правила модель НЕ проверила (нет годного вердикта после всех попыток).
    # Они не должны выглядеть пройденными — движок пометит их «НЕ ПРОВЕРЕНО».
    checked_indices = set()
    for v in verdicts:
        if not isinstance(v, dict):
            continue
        ri = v.get("rule_index")
        verd = v.get("verdict")
        status = verd.get("status") if isinstance(verd, dict) else None
        has_status = isinstance(status, str) and status.strip() != ""
        if not has_status:
            continue
        if isinstance(ri, int):
            checked_indices.add(ri)
        elif isinstance(ri, str) and ri.strip().isdigit():
            checked_indices.add(int(ri.strip()))
    unchecked = [r for r in rules if r.get("index") not in checked_indices]
    if unchecked:
        logger.warning(
            f"multi_rule[{layer}]: не проверено {len(unchecked)}/{len(rules)} правил "
            f"(индексы {[r.get('index') for r in unchecked]}) — помечены «НЕ ПРОВЕРЕНО»"
        )
    if session_dir:
        (session_dir / f"{prefix}_response_raw.txt").write_text(response_text, encoding="utf-8")
        (session_dir / f"{prefix}_response_parsed.json").write_text(
            json.dumps(verdicts, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )

    violations = _verdict_to_violations(verdicts, rules, layer=layer)
    return {
        "verdicts": verdicts,
        "violations": violations,
        "unchecked": unchecked,
        "checked_count": len(rules) - len(unchecked),
        "usage": {
            "prompt_tokens": response.usage.prompt_tokens,
            "completion_tokens": response.usage.completion_tokens,
            "total_tokens": response.usage.total_tokens,
        },
        "prompt_chars": len(user_prompt),
    }
# END_RUN_MULTI_RULE_AUDIT


# START_CONFIG_LOADERS
# PURPOSE: Загрузчики `sections.json`, `rules_multi.json`, `rules_methodology.json` по doc_type.
# INPUTS: Путь к корню doc_configs, имя doc_type.
# OUTPUTS: Сконфигурированные dict'ы или None, если файлов нет (такой doc_type не использует multi-rule/methodology).
# KEYWORDS: config-loader, sections, rules, methodology.
def load_multi_rule_config(doc_configs_dir: Path, doc_type: str) -> Optional[Dict[str, Any]]:
    """
    Назначение:
        Загружает multi-rule конфигурацию (sections + rules_multi) для doc_type.

    Вход:
        doc_configs_dir: Корень `doc_configs/`.
        doc_type: Имя типа документа.

    Выход:
        `{sections, rules, include_scopes}` или `None` если файлов нет.

    Логика:
        Ищет `sections.json` и `rules_multi.json` в `doc_configs/<doc_type>/`. Если
        оба есть — возвращает собранный конфиг; если хотя бы одного нет — возвращает
        `None` (doc_type не поддерживает multi-rule).
    """
    base = Path(doc_configs_dir) / doc_type
    sections_path = base / "sections.json"
    rules_path = base / "rules_multi.json"
    if not sections_path.exists() or not rules_path.exists():
        return None
    sections_data = json.loads(sections_path.read_text(encoding="utf-8"))
    rules_data = json.loads(rules_path.read_text(encoding="utf-8"))
    rules = list(rules_data["rules"])
    # Overlay: пользовательские правила (rules_custom.json) добавляются к базовому слою,
    # если файл есть. Индексы кастома — с 200 (не пересекаются с base 1-99). Читается
    # заново на каждый аудит → правки в UI применяются без рестарта сервиса.
    custom_path = base / "rules_custom.json"
    warnings: List[str] = []
    if custom_path.exists():
        try:
            custom_data = json.loads(custom_path.read_text(encoding="utf-8"))
            rules.extend(custom_data.get("rules", []))
        except (json.JSONDecodeError, KeyError) as e:
            # F4: битый файл пользовательских правил — проверка идёт без них, но с предупреждением в результате.
            warnings.append(f"Пользовательские правила не применены: файл rules_custom.json повреждён ({e}). Обратитесь к администратору.")
    return {
        "sections": sections_data["sections"],
        "rules": rules,
        "include_scopes": rules_data.get("include_scopes"),
        "warnings": warnings,
    }


def load_methodology_config(doc_configs_dir: Path, doc_type: str) -> Optional[Dict[str, Any]]:
    """
    Назначение:
        Загружает методический конфиг (rules_methodology.json) для doc_type.

    Вход:
        doc_configs_dir: Корень `doc_configs/`.
        doc_type: Имя типа документа.

    Выход:
        `{rules, include_scopes, source}` или `None` если файла нет.

    Логика:
        Sections не читает — переиспользует ту же семантическую карту из `sections.json`
        (вызывающий код берёт sections из `load_multi_rule_config`).
    """
    base = Path(doc_configs_dir) / doc_type
    rules_path = base / "rules_methodology.json"
    if not rules_path.exists():
        return None
    rules_data = json.loads(rules_path.read_text(encoding="utf-8"))
    return {
        "rules": rules_data["rules"],
        "include_scopes": rules_data.get("include_scopes"),
        "source": rules_data.get("source", ""),
    }
# END_CONFIG_LOADERS
