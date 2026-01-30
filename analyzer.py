"""Анализ драйверов производительности с помощью LLM и генерация отчетов."""

import json
import re
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Callable, Dict, Iterable, List, Optional, Tuple

from openai import OpenAI
from openpyxl import Workbook


@dataclass
class DriverEntry:
    number: str
    driver_name: str
    average_score: Optional[float]
    problem_comment: str
    notes: List[Dict[str, str]]


DEVELOPER_INSTRUCTION = """
Ты — аналитик, проверяющий соответствие выводов драйверам блока. Тебе дают JSON со структурой:
- section_title
- score_thresholds (primary, fallback)
- selected_threshold
- eligible_drivers (номер, название, средняя оценка, комментарии)
- all_drivers (для контекста)
- summary_text
- summary_driver_map (словарь «имя драйвера -> текст вывода»)

Требования:
1. Проверь, присутствует ли каждый драйвер из eligible_drivers в summary_text и/или summary_driver_map.
    Сопоставляй по номеру (например "1)" или "-1") и по имени драйвера.
2. Для каждого драйвера сформируй объект вида:
    {
    "number": "...",
    "driver_name": "...",
    "average_score": ...,
    "found_in_summary": true/false,
    "evidence": "короткое пояснение или цитата из summary_text"
    }
3. Если драйвер отсутствует в выводах, добавь запись в remarks:
    {"number": "...", "driver_name": "...", "issue": "нет в выводах"}
4. Явно укажи, какой порог применился (selected_threshold).
5. Верни строго JSON без дополнительного текста:
{
"section_title": "...",
"applied_threshold": ...,
"driver_checks": [...],
"remarks": [...],
"summary_driver_map": {...}
}
"""


def parse_average(value: Optional[str]) -> Optional[float]:
    """Парсит среднюю оценку из строкового значения."""
    if value is None:
        return None
    cleaned = value.replace(",", ".").strip()
    if not cleaned:
        return None
    try:
        return float(cleaned)
    except ValueError:
        return None


def extract_summary_driver_map(summary_text: str) -> Dict[str, str]:
    """
    Извлекает карту драйверов из текста summary.

    Ожидает формат:
    1) Название драйвера: описание
    2) Другой драйвер: описание
    """
    blocks: Dict[str, str] = {}
    current_key: Optional[str] = None
    current_lines: List[str] = []

    bullet_pattern = re.compile(r"^\s*(\d+)\)\s*(.*)$")
    code_prefix = re.compile(r"^[-–]?\s*\d+[?.]?\s*")

    for raw_line in summary_text.splitlines():
        line = raw_line.strip()
        bullet = bullet_pattern.match(line)
        if bullet:
            if current_key is not None:
                blocks[current_key] = " ".join(filter(None, current_lines)).strip()
            remainder = bullet.group(2).strip()
            remainder = code_prefix.sub("", remainder, count=1).strip()
            current_key = remainder.split(":", 1)[0].strip() if remainder else f"Driver{bullet.group(1)}"
            first_text = remainder.split(":", 1)[1].strip() if ":" in remainder else ""
            current_lines = [first_text] if first_text else []
        elif current_key is not None and line:
            current_lines.append(line)
    if current_key is not None:
        blocks[current_key] = " ".join(filter(None, current_lines)).strip()
    return blocks


def to_driver_entries(section: Dict[str, object]) -> List[DriverEntry]:
    """Преобразует вопросы секции в список DriverEntry."""
    entries: List[DriverEntry] = []
    for q in section.get("questions", []):
        avg = parse_average(q["scores"].get("Средняя"))
        entries.append(
            DriverEntry(
                number=q.get("number", "").strip(),
                driver_name=q.get("question", "").strip(),
                average_score=avg,
                problem_comment=q.get("problem_comment", "").strip(),
                notes=q.get("notes", []),
            )
        )
    return entries


def prepare_section_context(
    section: Dict[str, object],
    primary_threshold: float = 9.0,
    fallback_threshold: float = 7.0,
) -> Dict[str, object]:
    """
    Подготавливает контекст секции для LLM-анализа.

    Фильтрует драйверы по порогам оценок и извлекает summary.
    """
    drivers = to_driver_entries(section)
    primary = [d for d in drivers if d.average_score is not None and d.average_score >= primary_threshold]
    selected_threshold = primary_threshold
    selected = primary
    if not selected:
        selected = [d for d in drivers if d.average_score is not None and d.average_score >= fallback_threshold]
        selected_threshold = fallback_threshold

    summary_text = section.get("summary", "") or ""
    summary_driver_map = extract_summary_driver_map(summary_text)

    return {
        "section_title": section.get("title"),
        "score_thresholds": {"primary": primary_threshold, "fallback": fallback_threshold},
        "selected_threshold": selected_threshold,
        "eligible_drivers": [
            {
                "number": d.number,
                "driver_name": d.driver_name,
                "average_score": d.average_score,
                "problem_comment": d.problem_comment,
                "notes": d.notes,
            }
            for d in selected
        ],
        "all_drivers": [
            {
                "number": d.number,
                "driver_name": d.driver_name,
                "average_score": d.average_score,
                "problem_comment": d.problem_comment,
                "notes": d.notes,
            }
            for d in drivers
        ],
        "summary_text": summary_text,
        "summary_driver_map": summary_driver_map,
    }


def call_driver_llm(
    client: OpenAI,
    section_payload: Dict[str, object],
    model: str = "gpt-4o-mini",
    temperature: float = 0,
) -> Tuple[str, Any]:
    """
    Вызывает LLM для анализа секции.

    Returns
    -------
    Tuple[str, Any]
        (JSON-текст результата, полный ответ API)
    """
    user_payload = json.dumps(section_payload, ensure_ascii=False, indent=2)
    response = client.chat.completions.create(
        model=model,
        temperature=temperature,
        messages=[
            {"role": "system", "content": DEVELOPER_INSTRUCTION},
            {"role": "user", "content": user_payload},
        ],
    )
    return response.choices[0].message.content, response


def collect_remarks_and_summaries(section_jsons: Iterable[Dict[str, Any]]) -> Tuple[List[Dict[str, Any]], Dict[str, str]]:
    """
    Собирает замечания и итоговые тексты драйверов из списка JSON-ответов LLM.

    Returns
    -------
    Tuple[List[Dict[str, Any]], Dict[str, str]]
        (remarks, driver_summary_map)
    """
    aggregated_remarks: List[Dict[str, Any]] = []
    driver_summary_map: Dict[str, str] = {}

    for section in section_jsons:
        section_title = section.get("section_title")
        for remark in section.get("remarks") or []:
            merged = dict(remark)
            if section_title and "section_title" not in merged:
                merged["section_title"] = section_title
            aggregated_remarks.append(merged)

        for driver_name, summary_text in (section.get("summary_driver_map") or {}).items():
            normalized_name = driver_name.strip()
            if not normalized_name:
                continue
            summary_text = (summary_text or "").strip()
            existing = driver_summary_map.get(normalized_name)
            if existing and existing != summary_text:
                driver_summary_map[normalized_name] = f"{existing}\n---\n{summary_text}"
            else:
                driver_summary_map[normalized_name] = summary_text

    return aggregated_remarks, driver_summary_map


def build_question_lookup(parsed_sections: Iterable[Dict[str, Any]]) -> Dict[Tuple[str, str], Dict[str, Any]]:
    """
    Создаёт справочник {(название блока, номер) -> данные по вопросу}.

    Помогает восстановить исходные комментарии/заметки из распарсенной таблицы.
    """
    lookup: Dict[Tuple[str, str], Dict[str, Any]] = {}
    for section in parsed_sections:
        title = (section.get("title") or "").strip()
        for question in section.get("questions", []):
            key = (title, (question.get("number") or "").strip())
            lookup[key] = {
                "driver_name": (question.get("question") or "").strip(),
                "problem_comment": (question.get("problem_comment") or "").strip(),
                "notes": "\n".join(note.get("text", "").strip() for note in question.get("notes", [])),
            }
    return lookup


def export_missing_driver_report(
    parsed_data: Dict[str, Any],
    section_results: Iterable[Dict[str, Any]],
    output_path: Path,
    sheet_name: str = "Итог",
) -> None:
    """
    Формирует Excel-отчёт с драйверами, которые не нашли в выводах (summary).

    Parameters
    ----------
    parsed_data : Dict[str, Any]
        Словарь, загруженный из результата парсинга
    section_results : Iterable[Dict[str, Any]]
        Итерируемый набор JSON-ответов LLM
    output_path : Path
        Куда сохранить Excel
    sheet_name : str
        Имя листа (по умолчанию "Итог")
    """
    question_lookup = build_question_lookup(parsed_data.get("sections", []))

    wb = Workbook()
    ws = wb.active
    ws.title = sheet_name

    headers = [
        "Блок",
        "Номер драйвера",
        "Наименование драйвера",
        "Средняя оценка",
        "Комментарий (проблема)",
        "Дополнительные примечания",
        "Замечание",
    ]
    ws.append(headers)

    for section in section_results:
        title = (section.get("section_title") or "").strip()
        remark_lookup = {
            (remark.get("number"), remark.get("driver_name")): remark for remark in section.get("remarks") or []
        }

        for check in section.get("driver_checks") or []:
            if check.get("found_in_summary"):
                continue

            number = (check.get("number") or "").strip()
            driver_name = (check.get("driver_name") or "").strip()
            key = (title, number)
            parsed_info = question_lookup.get(key, {})
            remark = remark_lookup.get((number, driver_name), {})

            ws.append(
                [
                    title,
                    number,
                    driver_name or parsed_info.get("driver_name", ""),
                    check.get("average_score"),
                    parsed_info.get("problem_comment", ""),
                    parsed_info.get("notes", ""),
                    remark.get("issue", "Нет в выводах"),
                ]
            )

    wb.save(output_path)


def analyze_sections(
    parsed_data: Dict[str, Any],
    client: OpenAI,
    primary_threshold: float = 9.0,
    fallback_threshold: float = 7.0,
    model: str = "gpt-4o-mini",
    temperature: float = 0,
    progress_callback: Optional[Callable[[str, int, int], None]] = None,
) -> List[Dict[str, Any]]:
    """
    Анализирует все секции с помощью LLM.

    Parameters
    ----------
    parsed_data : Dict[str, Any]
        Распарсенные данные из Excel
    client : OpenAI
        Клиент OpenAI API
    primary_threshold : float
        Основной порог оценки
    fallback_threshold : float
        Запасной порог
    model : str
        Модель GPT
    temperature : float
        Температура генерации
    progress_callback : Optional[Callable[[str, int, int], None]]
        Callback функция для отслеживания прогресса (message, current, total)

    Returns
    -------
    List[Dict[str, Any]]
        Список результатов анализа по каждой секции
    """
    section_results = []
    total_sections = len(parsed_data["sections"])

    for idx, section in enumerate(parsed_data["sections"], 1):
        if progress_callback:
            section_title = section.get("title", f"Секция {idx}")
            progress_callback(f"Анализ: {section_title}", idx, total_sections)

        payload = prepare_section_context(section, primary_threshold, fallback_threshold)
        answer_text, _ = call_driver_llm(client, payload, model, temperature)
        section_json = json.loads(answer_text)
        section_results.append(section_json)

    return section_results
