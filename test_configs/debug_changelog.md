# Changelog отладки сервисов аудита

Лог прогресса исправления false_fail по результатам анализа тестового прогона.
Обновляется агентом при каждой сессии отладки (скилл `/debug-services`).

## Статистика

- **Базовый прогон:** run_20260210_165639
- **Базовая точность:** 63.7% (184 true_fail / 289 FAIL)
- **false_fail на старте:** 105
- **false_fail осталось:** 105
- **Текущая точность:** 63.7%

## Инфраструктура

- [x] compare_runs.py — скрипт сравнения прогонов (2026-02-12)
- [x] tests/ — директория юнит-тестов + conftest.py + test_filename_checker.py (15 тестов PASS) (2026-02-12)
- [x] filename_keywords — поддержка в check_filename_universal() + 4 конфига обновлены (2026-02-12)

## Прогресс по doc_type

| # | doc_type | false_fail | Статус | Новая точность |
|---|----------|------------|--------|----------------|
| 1 | kpsc | 27 | in_progress | 40.0% → ожидается ~85%+ |
| 2 | prikaz_comp_ppu | 10 | pending | 52.4% |
| 3 | polozhenie_ppu | 11 | pending | 42.1% |
| 4 | polozhenie_comp_ppu | 6 | pending | 60.0% |
| 5 | prikaz_ppu | 3 | pending | 72.7% |
| 6 | prikaz_ic_potoka | 8 | pending | 71.4% |
| 7 | prikaz_vyhod | 8 | pending | 71.4% |
| 8 | cheklist_eu | 7 | pending | 58.8% |
| 9 | prikaz_formirovanie_po | 5 | pending | 81.5% |
| 10 | prikaz_ic | 7 | pending | 78.1% |
| 11 | polozhenie_po | 6 | pending | 76.9% |
| 12 | presentation_eu | 2 | pending | 75.0% |

## Сессии отладки

<!-- Формат записи:

## [ГГГГ-ММ-ДД] — doc_type

### Что исправлено
- Root cause X: описание фикса (файл:строка)

### Результат перезапуска
- Прогон: run_YYYYMMDD_HHMMSS
- compare_runs.py:
  - ИСПРАВЛЕНО (FAIL → PASS): N
  - РЕГРЕССИИ (PASS → FAIL): N
  - БЕЗ ИЗМЕНЕНИЙ: N
- false_fail: было X → стало Y
- Точность: была X% → стала Y%

### Верификация
- Новые FAIL проверены: да/нет
- Регрессии: описание (если есть)

### Файлы изменены
- path/to/file1.py
- path/to/file2.json
-->

## [2026-02-12] — kpsc

### Что исправлено

**Root cause A — все 27 false_fail связаны с ошибками парсинга:**

1. **sheet_finder.py (новый файл)**
   - Создан универсальный fuzzy-поиск листов с Latin→Cyrillic нормализацией
   - Поддержка: trailing spaces, Latin C vs Cyrillic С, суффиксы (ВПП, ТС), подстроки
   - Алгоритм: keywords matching → exclude → prefer → выбор листа с макс. данными

2. **parse_kpsc_header.py — динамический маппинг полей**
   - Был: захардкоженные координаты C4-C8 (работало только для одного формата)
   - Стало: поиск лейблов по ключевым словам ("Ответственный", "Дата разработки" и т.д.)
   - Поддержка inline-значений ("Поток: Название...")
   - Фикс REGION_MAX_COL=31 → динамический ws.max_column
   - Исправляет: #6, #7, #8, #10 (biznes_otel), #28-30, #32 (rotosnab)

3. **Все 9 парсеров переведены на sheet_finder:**
   - parse_kpsc_table1.py: КПСЦ-лист (mapper, sodex, rotosnab)
   - parse_legend.py: "Условные обозначения" (graceful empty для mapper, sodex)
   - parse_loss_digitization.py: "Оцифровка потерь" (rotosnab, ruslet; graceful для остальных)
   - parse_pa1_chart.py: "ПА-1" / "ПА1" / "ПА1 ВПП" (все 5 компаний)
   - parse_pa1_table.py: аналогично
   - parse_pokazateli.py: "Показатели" / "Расчет показателей" (sodex)
   - parse_spaghetti_sheet.py: "Диаграмма Спагетти" (Latin C нормализация)
   - parse_spaghetti_problems.py: "Перечень проблем" / "Пробл. и улучш."

4. **Фикс _sheet_name_from_range в parse_pa1_chart.py**
   - Ссылки `[2]ПА1 '!$A$1` из chart series — убираем `[N]` префикс

### Результат парсинга (parser_summary comparison)

| Компания | Было OK | Стало OK | Исправлено |
|----------|---------|----------|------------|
| biznes_otel | 4 | 9 | +5 парсеров |
| mapper | 1 | 9 | +8 парсеров |
| rotosnab | 4 | 9 | +5 парсеров |
| ruslet | N/A* | 9 | +9 парсеров |
| sodex | 1 | 9 | +8 парсеров |

*ruslet: старый прогон не имел parser_summary

### Header-поля исправлены

| Компания | Поля: было→стало |
|----------|-----------------|
| biznes_otel | flow_name: дата→"Предоставление услуг"; date_developed: None→2025-06-24; compiled_by: None→заполнено |
| mapper | title: None→заполнено; flow_name: None→заполнено; compiled_by: None→заполнено |
| rotosnab | лист КПСЦ→КПСЦ ТС; все 4 поля: None→заполнены |
| ruslet | все 5 полей: None→заполнены |
| sodex | title: None→заполнено; compiled_by: None→заполнено |

### Ожидаемые фиксы false_fail (27 записей)

- Правила 7.x (15 записей): файлы парсинга созданы → валидаторы получат данные → PASS
- Правила 1.x (8 записей): header-поля заполнены → LLM корректно валидирует → PASS
- Правила 2.1, 6.1, 8.2 (4 записи): данные парсинга доступны → зависит от содержимого

### Блокер полного прогона

OpenAI API квота исчерпана (`insufficient_quota`). Все 25 валидаторов возвращают ERROR.
Парсинг (9/9 парсеров × 5 компаний = 45/45 OK) верифицирован.
Полный прогон с compare_runs.py требует рабочий API ключ.

### Юнит-тесты: 35/35 PASS

```
pytest tests/test_kpsc_parsers.py -v
# TestSheetFinder: 17 tests
# TestParseKpscHeader: 8 tests
# TestAllParsersRun: 10 tests
```

### Файлы изменены

- `audit_engine/kpsc/sheet_finder.py` (новый) — нечёткий поиск листов
- `audit_engine/kpsc/parser_scripts/parse_kpsc_header.py` — динамический маппинг + sheet_finder
- `audit_engine/kpsc/parser_scripts/parse_kpsc_table1.py` — sheet_finder
- `audit_engine/kpsc/parser_scripts/parse_legend.py` — sheet_finder + graceful empty
- `audit_engine/kpsc/parser_scripts/parse_loss_digitization.py` — sheet_finder + graceful empty
- `audit_engine/kpsc/parser_scripts/parse_pa1_chart.py` — sheet_finder + [N] prefix fix
- `audit_engine/kpsc/parser_scripts/parse_pa1_table.py` — sheet_finder + graceful empty
- `audit_engine/kpsc/parser_scripts/parse_pokazateli.py` — sheet_finder + graceful empty
- `audit_engine/kpsc/parser_scripts/parse_spaghetti_sheet.py` — sheet_finder + graceful empty
- `audit_engine/kpsc/parser_scripts/parse_spaghetti_problems.py` — sheet_finder + graceful empty
- `tests/test_kpsc_parsers.py` (новый) — 35 юнит-тестов
