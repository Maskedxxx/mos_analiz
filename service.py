#!/usr/bin/env python3
"""Сервис анализа драйверов производительности."""

import argparse
import json
import os
import shutil
import sys
import uuid
from datetime import datetime
from pathlib import Path
from typing import Dict, Optional

import yaml
from openai import OpenAI

from analyzer import analyze_sections, collect_remarks_and_summaries, export_missing_driver_report
from parser import parse_excel_to_json


def load_config(config_path: Path) -> Dict:
    """Загружает конфигурацию из YAML файла."""
    with open(config_path, "r", encoding="utf-8") as f:
        config = yaml.safe_load(f)
    return config


def create_session_directory(results_dir: Path, input_filename: str) -> Path:
    """
    Создает уникальную папку для сессии.

    Формат: YYYY-MM-DD_HHMMSS_xxxxx
    где xxxxx - короткий UUID
    """
    timestamp = datetime.now().strftime("%Y-%m-%d_%H%M%S")
    short_uuid = str(uuid.uuid4())[:8]
    session_name = f"{timestamp}_{short_uuid}"

    session_dir = results_dir / session_name
    session_dir.mkdir(parents=True, exist_ok=True)

    return session_dir


def get_openai_client(config: Dict) -> OpenAI:
    """
    Создает клиента OpenAI.

    Приоритет: config['openai']['api_key'] -> OPENAI_API_KEY env
    """
    api_key = config.get("openai", {}).get("api_key")
    if not api_key:
        api_key = os.getenv("OPENAI_API_KEY")
    if not api_key:
        raise ValueError(
            "OpenAI API key not found. Set it in config.yaml or OPENAI_API_KEY environment variable."
        )
    return OpenAI(api_key=api_key)


def process_driver_file(
    input_excel: Path,
    config_path: Path = Path("config.yaml"),
    output_excel: Optional[Path] = None,
) -> Path:
    """
    Основная функция сервиса: обрабатывает Excel файл и генерирует отчет.

    Parameters
    ----------
    input_excel : Path
        Путь к входному Excel файлу
    config_path : Path
        Путь к конфигурационному файлу
    output_excel : Optional[Path]
        Путь к выходному Excel файлу (если None, генерируется автоматически)

    Returns
    -------
    Path
        Путь к папке сессии с результатами
    """
    # 1. Загрузка конфигурации
    print(f"📝 Загрузка конфигурации из {config_path}")
    config = load_config(config_path)

    # 2. Создание папки сессии
    results_dir = Path(config.get("output", {}).get("results_dir", "results"))
    session_dir = create_session_directory(results_dir, input_excel.name)
    print(f"📁 Создана папка сессии: {session_dir}")

    # 3. Копирование входного файла
    input_copy = session_dir / f"00_input_{input_excel.name}"
    shutil.copy2(input_excel, input_copy)
    print(f"📄 Входной файл скопирован: {input_copy.name}")

    # 4. Парсинг Excel -> JSON
    print("🔍 Парсинг Excel файла...")
    parsed_data = parse_excel_to_json(input_excel)
    parsed_json_path = session_dir / "01_parsed.json"
    parsed_json_path.write_text(json.dumps(parsed_data, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"✅ Результат парсинга сохранен: {parsed_json_path.name}")

    # 5. Анализ с помощью LLM
    print("🤖 Запуск анализа с помощью LLM...")
    client = get_openai_client(config)

    analysis_config = config.get("analysis", {})
    primary_threshold = analysis_config.get("primary_threshold", 9.0)
    fallback_threshold = analysis_config.get("fallback_threshold", 7.0)

    openai_config = config.get("openai", {})
    model = openai_config.get("model", "gpt-4o-mini")
    temperature = openai_config.get("temperature", 0)

    print(f"   Пороги оценок: primary={primary_threshold}, fallback={fallback_threshold}")
    print(f"   Модель: {model}")

    section_results = analyze_sections(
        parsed_data=parsed_data,
        client=client,
        primary_threshold=primary_threshold,
        fallback_threshold=fallback_threshold,
        model=model,
        temperature=temperature,
    )

    # 6. Агрегация результатов
    remarks, driver_summary_map = collect_remarks_and_summaries(section_results)

    result_bundle = {
        "section_results": section_results,
        "remarks": remarks,
        "driver_summary_map": driver_summary_map,
    }

    llm_json_path = session_dir / "02_llm_analysis.json"
    llm_json_path.write_text(json.dumps(result_bundle, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"✅ Результат LLM-анализа сохранен: {llm_json_path.name}")

    # 7. Генерация Excel отчета
    print("📊 Генерация Excel отчета...")
    if output_excel is None:
        output_excel = session_dir / "03_report.xlsx"
    else:
        output_excel = session_dir / f"03_{output_excel.name}"

    sheet_name = config.get("output", {}).get("sheet_name", "Итог")

    export_missing_driver_report(
        parsed_data=parsed_data,
        section_results=section_results,
        output_path=output_excel,
        sheet_name=sheet_name,
    )
    print(f"✅ Excel отчет сохранен: {output_excel.name}")

    # 8. Вывод итоговой информации
    print("\n" + "=" * 60)
    print("🎉 Обработка завершена успешно!")
    print(f"📁 Папка сессии: {session_dir.absolute()}")
    print(f"📊 Итоговый отчет: {output_excel.name}")
    print(f"🔍 Найдено замечаний: {len(remarks)}")
    print("=" * 60)

    return session_dir


def main(argv=None):
    """CLI интерфейс сервиса."""
    parser = argparse.ArgumentParser(
        description="Сервис анализа драйверов производительности",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Примеры использования:
  python service.py driver_1.xlsx
  python service.py driver_1.xlsx -o report.xlsx
  python service.py driver_1.xlsx -c custom_config.yaml

Все результаты сохраняются в папку results/<session_id>/
        """,
    )

    parser.add_argument(
        "input",
        type=Path,
        help="Путь к входному Excel файлу с драйверами",
    )

    parser.add_argument(
        "-o",
        "--output",
        type=Path,
        help="Имя выходного Excel файла (опционально)",
    )

    parser.add_argument(
        "-c",
        "--config",
        type=Path,
        default=Path("config.yaml"),
        help="Путь к файлу конфигурации (по умолчанию: config.yaml)",
    )

    args = parser.parse_args(argv)

    # Проверка существования файлов
    if not args.input.exists():
        print(f"❌ Ошибка: файл {args.input} не найден", file=sys.stderr)
        return 1

    if not args.config.exists():
        print(f"❌ Ошибка: файл конфигурации {args.config} не найден", file=sys.stderr)
        return 1

    try:
        process_driver_file(
            input_excel=args.input,
            config_path=args.config,
            output_excel=args.output,
        )
        return 0
    except Exception as e:
        print(f"\n❌ Ошибка при обработке: {e}", file=sys.stderr)
        import traceback

        traceback.print_exc()
        return 1


if __name__ == "__main__":
    sys.exit(main())
