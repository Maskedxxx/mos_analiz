# START_MODULE_CONTRACT
# PURPOSE: Аудит документа «Драйверы производительности» (xlsx-опросник). По каждой секции выбираются драйверы со средним баллом ≥ primary_threshold (или fallback_threshold), отправляются в LLM с промптом из `doc_configs/drivers/driver_check_prompt.txt`, ответы агрегируются в Excel-отчёт о пропущенных в саммари драйверах.
# INPUTS: xlsx-файл (через `parse_excel_to_json` из `src/doc_type_parsers/drivers.py`); `doc_configs/drivers/config.json` (model, thresholds); `doc_configs/drivers/driver_check_prompt.txt` (system prompt); LLM-сервис через `src.llm.client.call_llm`.
# OUTPUTS: `run_drivers_special(args) -> AuditResult` — публичная точка, регистрируется в SPECIAL_ENGINE_RUNNERS. Также пишет в session_dir: `01_parsed.json`, `02_llm_analysis.json`, `03_report.xlsx`.
# KEYWORDS: validators, drivers, llm, xlsx-questionnaire.
# LINKS: src/doc_type_parsers/drivers.py (parse_excel_to_json), src/llm/client.py (call_llm), config/llm.py (LLM_CONFIG), doc_configs/drivers/.
# RATIONALE:
#   Интерфейс отличается от kpsc/kartochka/plan_grafik: drivers — не набор
#   независимых правил, а одна композитная проверка (на каждую секцию опросника
#   LLM смотрит выбранные драйверы и пишет remarks). Раньше внутри был свой
#   `OpenAI(api_key, base_url)` клиент и raw `client.chat.completions.create` с
#   try/fallback на response_format — оба вызова теперь идут через единый
#   `call_llm`, fallback сохранён через try/except.
# END_MODULE_CONTRACT

from __future__ import annotations

# START_IMPORTS
import json
import logging
import os
import re
import time
import traceback
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Tuple

from openpyxl import Workbook

from config.llm import LLM_CONFIG
from src.doc_type_parsers.drivers import parse_excel_to_json
from src.llm.client import call_llm, resolve_runtime_llm_model
# END_IMPORTS


logger = logging.getLogger(__name__)


# START_PATHS
# PURPOSE: Локальные пути модуля. Разрешаются из `src/doc_type_validators/drivers.py`,
# поэтому `parents[2]` указывает на корень репо.
_REPO_ROOT = Path(__file__).resolve().parents[2]
_DOC_CONFIGS_DIR = _REPO_ROOT / "doc_configs"
_LOGS_RESULT_DIR = _REPO_ROOT / "logs_result"
# END_PATHS


# START_DATA
@dataclass
class DriverEntry:
    """Один драйвер из секции опросника (после парсинга xlsx)."""
    number: str
    driver_name: str
    average_score: Optional[float]
    problem_comment: str
    notes: List[Dict[str, str]]
# END_DATA


# START_PROMPT_LOADING
def _default_prompt_path() -> Path:
    """Дефолтный путь к driver_check_prompt.txt (в doc_configs/drivers/)."""
    return _DOC_CONFIGS_DIR / "drivers" / "driver_check_prompt.txt"


def load_driver_prompt() -> str:
    """Читает system-prompt для drivers. Env DRIVERS_PROMPT_PATH переопределяет путь."""
    env_path = os.environ.get("DRIVERS_PROMPT_PATH")
    path = Path(env_path) if env_path else _default_prompt_path()
    return path.read_text(encoding="utf-8")


def _load_drivers_config() -> Dict[str, Any]:
    """Читает doc_configs/drivers/config.json (model, thresholds)."""
    with open(_DOC_CONFIGS_DIR / "drivers" / "config.json", "r", encoding="utf-8") as f:
        return json.load(f)
# END_PROMPT_LOADING


# START_SECTION_PREPARATION
def parse_average(value: Optional[str]) -> Optional[float]:
    """Парсит строку среднего балла в float (учитывает русскую запятую)."""
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
    Назначение:
        Извлекает маппинг {driver_name: summary_text} из текста саммари секции.
        Саммари секции — нумерованный список («1) ... 2) ...»), каждый блок
        соответствует одному драйверу.

    Логика:
        Бьёт текст по регулярке `\\d+)`, отделяет имя драйвера (до двоеточия)
        от тела блока, агрегирует подстроки в один dict.
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
    """Преобразует dict секции (от парсера) в список DriverEntry."""
    entries: List[DriverEntry] = []
    for q in section.get("questions", []):
        avg = parse_average((q.get("scores") or {}).get("Средняя"))
        entries.append(DriverEntry(
            number=(q.get("number") or "").strip(),
            driver_name=(q.get("question") or "").strip(),
            average_score=avg,
            problem_comment=(q.get("problem_comment") or "").strip(),
            notes=q.get("notes", []),
        ))
    return entries


def prepare_section_context(
    section: Dict[str, object],
    primary_threshold: float = 9.0,
    fallback_threshold: float = 7.0,
) -> Dict[str, object]:
    """
    Собирает payload для LLM по одной секции:
    - eligible_drivers: те у кого score ≥ primary_threshold (или ≥ fallback_threshold если primary пусто).
    - all_drivers: для контекста.
    - summary_text + summary_driver_map: предзаполненное саммари секции.
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
            {"number": d.number, "driver_name": d.driver_name, "average_score": d.average_score,
             "problem_comment": d.problem_comment, "notes": d.notes}
            for d in selected
        ],
        "all_drivers": [
            {"number": d.number, "driver_name": d.driver_name, "average_score": d.average_score,
             "problem_comment": d.problem_comment, "notes": d.notes}
            for d in drivers
        ],
        "summary_text": summary_text,
        "summary_driver_map": summary_driver_map,
    }
# END_SECTION_PREPARATION


# START_LLM_CALL
def call_driver_llm(
    section_payload: Dict[str, object],
    model: str,
    base_url: str,
    temperature: float,
    system_prompt: str,
) -> str:
    """
    Назначение:
        Один LLM-вызов для одной секции опросника. Идёт через единый
        `src.llm.client.call_llm`. Сначала пробует `response_format={'type':'json_object'}`,
        при ошибке (модель не поддерживает) — повторяет без response_format.

    Вход:
        section_payload: dict-payload от `prepare_section_context` (eligible_drivers,
            summary, и т.д.) — будет упакован в JSON и передан как user-сообщение.
        model, base_url, temperature: параметры LLM.
        system_prompt: текст driver_check_prompt.txt.

    Выход:
        Текст ответа LLM (JSON-строка).
    """
    user_payload = json.dumps(section_payload, ensure_ascii=False, indent=2)
    messages = [
        {"role": "system", "content": system_prompt},
        {"role": "user", "content": user_payload},
    ]
    try:
        return call_llm(
            messages=messages, model=model, base_url=base_url,
            temperature=temperature, response_format={"type": "json_object"},
        )
    except Exception:
        return call_llm(
            messages=messages, model=model, base_url=base_url,
            temperature=temperature,
        )


def analyze_sections(
    parsed_data: Dict[str, Any],
    primary_threshold: float,
    fallback_threshold: float,
    model: str,
    base_url: str,
    temperature: float,
    progress_callback: Optional[callable] = None,
    system_prompt: Optional[str] = None,
) -> List[Dict[str, Any]]:
    """Прогоняет все секции `parsed_data` через `call_driver_llm`. Возвращает список JSON-ответов LLM по секциям."""
    prompt = system_prompt or load_driver_prompt()
    section_results: List[Dict[str, Any]] = []
    sections = parsed_data.get("sections") or []
    total_sections = len(sections)
    for idx, section in enumerate(sections, 1):
        if progress_callback:
            title = section.get("title") or f"Секция {idx}"
            progress_callback(f"Анализ: {title}", idx, total_sections)
        payload = prepare_section_context(section, primary_threshold, fallback_threshold)
        answer_text = call_driver_llm(payload, model, base_url, temperature, prompt)
        section_json = json.loads(answer_text)
        section_results.append(section_json)
    return section_results
# END_LLM_CALL


# START_AGGREGATION
def collect_remarks_and_summaries(section_jsons: Iterable[Dict[str, Any]]) -> Tuple[List[Dict[str, Any]], Dict[str, str]]:
    """Агрегирует remarks (нарушения) и summary_driver_map (саммари по драйверам) из всех секций."""
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
            normalized_name = (driver_name or "").strip()
            if not normalized_name:
                continue
            summary_text = (summary_text or "").strip()
            existing = driver_summary_map.get(normalized_name)
            if existing and existing != summary_text:
                driver_summary_map[normalized_name] = f"{existing}\n---\n{summary_text}"
            else:
                driver_summary_map[normalized_name] = summary_text
    return (aggregated_remarks, driver_summary_map)


def build_question_lookup(parsed_sections: Iterable[Dict[str, Any]]) -> Dict[Tuple[str, str], Dict[str, Any]]:
    """Маппинг (section_title, driver_number) → данные драйвера из парсера. Используется для обогащения Excel-отчёта."""
    lookup: Dict[Tuple[str, str], Dict[str, Any]] = {}
    for section in parsed_sections:
        title = (section.get("title") or "").strip()
        for question in section.get("questions", []):
            key = (title, (question.get("number") or "").strip())
            lookup[key] = {
                "driver_name": (question.get("question") or "").strip(),
                "problem_comment": (question.get("problem_comment") or "").strip(),
                "notes": "\n".join((note.get("text", "").strip() for note in question.get("notes", []))),
            }
    return lookup


def export_missing_driver_report(
    parsed_data: Dict[str, Any],
    section_results: Iterable[Dict[str, Any]],
    output_path: Path,
    sheet_name: str = "Итог",
) -> None:
    """Пишет Excel-отчёт о драйверах, которые не попали в summary секции (т.е. пропущены в выводах опросника)."""
    question_lookup = build_question_lookup(parsed_data.get("sections", []))
    wb = Workbook()
    ws = wb.active
    ws.title = sheet_name
    ws.append(["Блок", "Номер драйвера", "Наименование драйвера", "Средняя оценка",
               "Комментарий (проблема)", "Дополнительные примечания", "Замечание"])
    for section in section_results:
        title = (section.get("section_title") or "").strip()
        remark_lookup = {(remark.get("number"), remark.get("driver_name")): remark for remark in section.get("remarks") or []}
        for check in section.get("driver_checks") or []:
            if check.get("found_in_summary"):
                continue
            number = (check.get("number") or "").strip()
            driver_name = (check.get("driver_name") or "").strip()
            key = (title, number)
            parsed_info = question_lookup.get(key, {})
            remark = remark_lookup.get((number, driver_name), {})
            ws.append([title, number, driver_name or parsed_info.get("driver_name", ""),
                       check.get("average_score"), parsed_info.get("problem_comment", ""),
                       parsed_info.get("notes", ""), remark.get("issue", "Нет в выводах")])
    output_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output_path)
# END_AGGREGATION


# START_PIPELINE
def _run_drivers_pipeline(
    args,
    config: Dict[str, Any],
    model: str,
    temperature: float,
    primary_threshold: float,
    fallback_threshold: float,
    session_dir: Path,
    start_time: float,
):
    """Основной пайплайн drivers (4 шага). Возвращает AuditResult — типизирован Any во избежание циклического импорта."""
    from main import AuditResult  # late import — main.py импортирует этот модуль, поэтому здесь lazy
    target_path = Path(args.target)
    print("[1/4] Парсинг Excel...")
    parsed = parse_excel_to_json(target_path)
    parsed_path = session_dir / "01_parsed.json"
    with open(parsed_path, "w", encoding="utf-8") as f:
        json.dump(parsed, f, ensure_ascii=False, indent=2)
    print(f"  Секций: {len(parsed.get('sections', []))}")
    print(f"  Сохранено: {parsed_path}")
    if args.parse_only:
        print(f"\n{'=' * 60}")
        print("TARGET (parsed):")
        print(json.dumps(parsed, ensure_ascii=False, indent=2))
        return AuditResult(doc_type="drivers", session_dir=session_dir, target_path=str(target_path), duration_sec=time.time() - start_time)

    print("\n[2/4] LLM-анализ секций...")
    base_url = config.get("llm_base_url", LLM_CONFIG.base_url)
    model = resolve_runtime_llm_model(model)
    system_prompt = load_driver_prompt()

    def progress_cb(msg: str, idx: int, total: int):
        print(f"  [{idx}/{total}] {msg}")

    section_results = analyze_sections(
        parsed_data=parsed,
        primary_threshold=primary_threshold,
        fallback_threshold=fallback_threshold,
        model=model,
        base_url=base_url,
        temperature=temperature,
        progress_callback=progress_cb,
        system_prompt=system_prompt,
    )
    analysis_path = session_dir / "02_llm_analysis.json"
    with open(analysis_path, "w", encoding="utf-8") as f:
        json.dump(section_results, f, ensure_ascii=False, indent=2)
    print(f"  Сохранено: {analysis_path}")

    print("\n[3/4] Сбор замечаний...")
    remarks, _driver_summary_map = collect_remarks_and_summaries(section_results)
    print(f"  Замечаний: {len(remarks)}")

    print("\n[4/4] Генерация Excel-отчёта...")
    report_path = session_dir / "03_report.xlsx"
    export_missing_driver_report(parsed, section_results, report_path)
    print(f"  Отчёт: {report_path}")

    duration = time.time() - start_time
    violations = [
        {
            "rule_index": f"driver_{r.get('section_title', 'x')}_{r.get('number', '?')}",
            "rule_title": r.get("driver_name", ""),
            "section_title": r.get("section_title", ""),
            "issue": r.get("issue", ""),
        }
        for r in remarks
    ]
    print(f"\n{'=' * 60}")
    print(f"  Итого нарушений: {len(violations)}")
    if violations:
        for v in violations:
            print(f"  - [{v['rule_index']}] {v['rule_title']}: {v['issue']}")
    print(f"  Время: {duration:.1f} сек")
    print(f"  Сессия: {session_dir}")
    print(f"{'=' * 60}")
    return AuditResult(
        violations=violations, doc_type="drivers", session_dir=session_dir,
        duration_sec=duration, rules_checked=len(parsed.get("sections", [])),
        target_path=str(target_path),
    )


def run_drivers_special(args):
    """
    Назначение:
        Публичный entrypoint для drivers — регистрируется в `SPECIAL_ENGINE_RUNNERS`
        в main.py. Подгружает config, разрешает model/temperature/thresholds,
        создаёт session_dir, запускает пайплайн.

    Вход:
        args: argparse-Namespace c полями {target, model?, temperature?, parse_only?, session_dir?}.

    Выход:
        AuditResult.
    """
    start_time = time.time()
    config = _load_drivers_config()
    model = args.model or config.get("model", LLM_CONFIG.default_model)
    temperature = args.temperature if args.temperature is not None else config.get("temperature", 0.0)
    primary_threshold = config.get("primary_threshold", 9.0)
    fallback_threshold = config.get("fallback_threshold", 7.0)
    session_dir = (
        Path(args.session_dir) if args.session_dir
        else _LOGS_RESULT_DIR / "drivers" / f"session_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
    )
    session_dir.mkdir(parents=True, exist_ok=True)
    try:
        return _run_drivers_pipeline(args, config, model, temperature, primary_threshold, fallback_threshold, session_dir, start_time)
    except Exception as e:
        error_msg = f"{type(e).__name__}: {e}\n{traceback.format_exc()}"
        error_path = session_dir / "ERROR.txt"
        error_path.write_text(error_msg, encoding="utf-8")
        logger.error("drivers: ошибка аудита → %s: %s", error_path, e)
        print(f"\n  ОШИБКА: {e}")
        print(f"  Лог ошибки: {error_path}")
        raise
# END_PIPELINE
