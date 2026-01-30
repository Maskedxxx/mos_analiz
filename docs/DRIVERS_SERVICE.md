# Сервис проверки драйверов (Excel) — карта кода и принцип работы

Обновлено: 2026-01-16

Цель сервиса: автоматически проверить, что **высокооценённые драйверы** из Excel отражены в **итоговых выводах (summary)**, и сформировать Excel‑отчёт по пропускам.

---

## 1) Что считать “сервисом драйверов” в этом репозитории

Ключевой код находится **в корне** репозитория:

- `README_SERVICE.md` — пользовательская документация (CLI + Web).
- `service.py` — CLI-оркестратор пайплайна (один файл → один прогон → папка результатов).
- `web_service.py` — веб‑обёртка FastAPI (upload + прогресс + скачивание отчёта).
- `parser.py` — парсер входного `.xlsx` в структурированный JSON (без pandas).
- `analyzer.py` — LLM-анализ секций + сбор замечаний + генерация Excel‑отчёта.
- `config.yaml` — пороги, модель, температура, выходная папка.
- `requirements.txt` — зависимости для CLI/Web сервиса.
- `templates/index.html` — UI веб‑страницы (Drag&Drop загрузка).
- `scripts/check_all_xlsx.py` — утилита для пакетной проверки корректности парсинга `.xlsx`.

Сгенерированные/вспомогательные папки (не “исходники” сервиса):
- `results/` — результаты запусков (сессии).
- `uploads/` — временные файлы веб‑загрузок.

---

## 2) Пайплайн (как работает end-to-end)

### CLI (`service.py`)

Основная функция: `process_driver_file()` в `service.py`.

Последовательность:
1) Загружает `config.yaml` (`load_config()`).
2) Создаёт папку сессии `results/<timestamp>_<uuid>/` (`create_session_directory()`).
3) Копирует входной `.xlsx` в сессию (`00_input_<file>.xlsx`).
4) Парсит Excel → JSON: `parser.parse_excel_to_json()` → сохраняет `01_parsed.json`.
5) Вызывает LLM по секциям: `analyzer.analyze_sections()` → сохраняет `02_llm_analysis.json`.
6) Генерирует Excel‑отчёт: `analyzer.export_missing_driver_report()` → `03_report.xlsx`.

### Web (`web_service.py`)

Веб‑версия делает тот же пайплайн, но:
- принимает файл через `POST /upload`,
- шлёт прогресс через WebSocket `/ws/{session_id}`,
- отдаёт готовые файлы через `GET /download/{session_id}/report` и `GET /download/{session_id}/json`.

---

## 3) Модули и ответственность (кто за что отвечает)

### `parser.py` — “из Excel сделать структуру”

Задача: превратить `.xlsx` в JSON со смысловыми блоками:
- `meta` (всё, что до первой секции),
- `sections[]` (каждая секция: заголовок, вопросы, summary).

Ключевые идеи парсинга:
- Чтение “как zip”: `.xlsx` открывается как архив, читаются `workbook.xml`, `sheet*.xml`, `sharedStrings.xml`.
- Выбор листа: по `sheet_name`, иначе — первый **не hidden**.
- Разметка секций:
  - строка, где `E == "Баллы"` — старт секции (в `B` ожидается заголовок блока).
  - “summary” определяется по merged‑диапазонам: если есть merge с началом в колонке `B` и шириной минимум до `J`, то эта строка считается строкой вывода (summary).
- Разметка вопросов:
  - номер вопроса берётся из колонки `B`, если значение “похоже на вопрос” (начинается с цифры).
  - текст вопроса — `C`, комментарий проблемы — `D`.
  - оценки — начиная с колонок `E+`: определяется layout (кто “респонденты”, где “Итого/Всего”, где “Средняя”).
  - строки‑примечания добавляются в `notes` к предыдущему вопросу.

Точка входа:
- `parse_excel_to_json(excel_path, sheet_name=None) -> dict`.

### `analyzer.py` — “из структуры сделать проверку + отчёт”

Задачи:
- отобрать драйверы по порогам,
- подготовить промпт/контекст секции,
- вызвать LLM,
- собрать нарушения,
- сформировать Excel‑таблицу “пропущенных драйверов”.

Ключевые функции:
- `prepare_section_context(section, primary_threshold, fallback_threshold)`:
  - строит `eligible_drivers`: сначала по `primary_threshold`, если пусто — по `fallback_threshold`;
  - кладёт `summary_text` и извлечённую “карту драйверов” `summary_driver_map`.
- `call_driver_llm(client, section_payload, model, temperature)`:
  - один запрос в LLM на одну секцию,
  - системная инструкция: `DEVELOPER_INSTRUCTION` (ответ строго JSON).
- `analyze_sections(parsed_data, client, ...)`:
  - идёт по `parsed_data["sections"]`,
  - вызывает LLM и парсит JSON ответ.
- `collect_remarks_and_summaries(section_jsons)`:
  - агрегирует `remarks` и объединяет `summary_driver_map` по секциям.
- `export_missing_driver_report(parsed_data, section_results, output_path, sheet_name)`:
  - строит Excel, добавляет строки только для `found_in_summary=false`.

### `service.py` — “запуск и артефакты”

Отвечает за:
- CLI аргументы (`input`, `--output`, `--config`);
- создание папки сессии;
- сохранение `00/01/02/03_*` файлов;
- получение OpenAI ключа (`OPENAI_API_KEY` или `config.yaml`).

### `web_service.py` — “упаковка в FastAPI”

Отвечает за:
- upload файла (`/upload`) и запуск фоновой обработки;
- прогресс по WebSocket (`ConnectionManager`);
- отдачу результатов и статуса (`/download/...`, `/status/...`).

---

## 4) Форматы входа/выхода (что лежит в `results/<session>/`)

Ожидаемый набор:
- `00_input_<name>.xlsx` — копия входного файла.
- `01_parsed.json` — результат `parser.py`.
- `02_llm_analysis.json` — результат `analyzer.py` (секции + remarks + aggregated map).
- `03_report.xlsx` — Excel‑отчёт: только драйверы, которые не найдены в выводах.

---

## 5) Конфигурация и окружение

### `config.yaml`

- `analysis.primary_threshold`, `analysis.fallback_threshold` — пороги отбора.
- `openai.model`, `openai.temperature` — параметры вызова модели.
- `openai.api_key` — может быть `null`; иначе используется `OPENAI_API_KEY` из окружения.
- `output.sheet_name`, `output.results_dir` — настройки отчёта/папки.

### Важно про ключ

В репозитории **не должно** быть реального API‑ключа: используйте `OPENAI_API_KEY` и/или хранение секретов вне git.

