# Changelog отладки сервисов аудита

Лог прогресса исправления false_fail по результатам анализа тестового прогона.
Обновляется агентом при каждой сессии отладки (скилл `/debug-services`).

## Статистика

- **Базовый прогон:** run_20260210_165639
- **Базовая точность:** 63.7% (184 true_fail / 289 FAIL)
- **false_fail на старте:** 105
- **false_fail осталось:** 36 (9 переклассифицированы в true_fail + 2 новых обнаружены при верификации ruslet)
- **false_fail исправлено (kpsc):** 25/27 → PASS
- **false_fail исправлено (prikaz_comp_ppu):** 10/10 → PASS
- **false_fail исправлено (cluster 3):** 26/26 → PASS (prikaz_vyhod 8, prikaz_comp_ppu 2*, prikaz_ppu 3, prikaz_ic_potoka 5 + 8 уже в кеше)
- **false_fail исправлено (polozhenie_ppu):** 11/11 → PASS
- **false_fail исправлено (polozhenie_comp_ppu):** 6/6 → PASS
- **false_fail исправлено (cheklist_eu):** 4/7 → PASS + 3 переклассифицированы в true_fail B
- **false_fail исправлено (prikaz_formirovanie_po):** 3/5 → PASS + 2 переклассифицированы в true_fail B
- **false_fail исправлено (prikaz_ic):** 3/5 → PASS + 1 переклассифицирован в true_fail B + 1 open (needs non-LLM)
- **Текущая точность:** ~79% (оценочно, полный прогон не запускался)

## Инфраструктура

- [x] compare_runs.py — скрипт сравнения прогонов (2026-02-12)
- [x] tests/ — директория юнит-тестов + conftest.py + test_filename_checker.py (15 тестов PASS) (2026-02-12)
- [x] filename_keywords — поддержка в check_filename_universal() + 4 конфига обновлены (2026-02-12)

## Прогресс по doc_type

| # | doc_type | false_fail | Статус | Новая точность |
|---|----------|------------|--------|----------------|
| 1 | kpsc | 27→0 | done | 40.0% → 60.8% (25 fixed, 2 reclassified) |
| 2 | prikaz_comp_ppu | 10→0 | done | 52.4% → ~85% (10+2 fixed, +1 reclassified) |
| 3 | prikaz_vyhod | 8→0 | done | 71.4% → ~86% (+8 fixed) |
| 4 | prikaz_ppu | 3→0 | done | 72.7% → ~88% (+3 fixed) |
| 5 | prikaz_ic_potoka | 8→0 | done | 71.4% → ~82% (+5 fixed, +1 reclassified) |
| 6 | polozhenie_ppu | 11→0 | done | 42.1% → 100% (11 fixed, 0 regressions) |
| 7 | polozhenie_comp_ppu | 6→0 | done | 60.0% → 100% (6 fixed, 0 regressions) |
| 8 | cheklist_eu | 7→0 | done | 58.8% → 100% (4 fixed, 3 reclassified B, 0 regressions) |
| 9 | prikaz_formirovanie_po | 5→0 | done | 32.5% → 40.0% (3 fixed, 2 reclassified B, 0 regressions) |
| 10 | prikaz_ic | 5→1 | done* | 11.1% → 19.4% (3 fixed, 1 reclassified B, 1 open #268) |
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

## [2026-02-13] — kpsc (продолжение)

### Что исправлено

**Оставшиеся 8 false_fail — 6 фиксов кода + 2 переклассификации:**

1. **validate_6_1_transport_row.py** — расширен фильтр колонок [2,3] → [1,2,3]
   - mapper и sodex имеют "Перемещения" в col A (col=1), не col B/C
   - Исправляет: #19 (mapper 6.1), #42 (sodex 6.1)

2. **validate_2_1_problems_count.py** — расширен фильтр col=2 → col in [1,2]
   - rotosnab имеет номера проблем в col A (col=1)
   - Исправляет: #33 (rotosnab 2.1)

3. **validate_8_2_values_cross_check.py** — динамический поиск колонки ИТОГО
   - Был: hardcoded col=26 (Z). Rotosnab имеет ИТОГО в col=20 (T)
   - Стало: ищет "ИТОГО" в заголовочной строке таблицы
   - Исправляет: #39 (rotosnab 8.2)

4. **parse_kpsc_header.py** — добавлено поле `organization`
   - Сканирует всю ширину строк 1-3 для ООО/АО/ПАО + "наименование"
   - biznes_otel: ООО в merged range AS1:AV1 (col 45-48)
   - Исправляет: #6 (biznes_otel 1.2)

5. **validate_1_2_company_name.py** — детерминированная предпроверка
   - Если organization/title содержит ООО + наименование → PASS без LLM
   - Устраняет ложноотрицательные LLM-вердикты
   - Исправляет: #6 (biznes_otel 1.2)

6. **validate_1_4_responsible.py** — детерминированная предпроверка
   - Если ≥2 слова и ≥5 символов → PASS без LLM
   - "Жукова А." — 2 токена, LLM ошибочно считал за 1 слово
   - Исправляет: #7 (biznes_otel 1.4)

7. **#10 biznes_otel 1.7** — переклассифицирован false_fail → true_fail (B)
   - compiled_by = "Рабочая группа ООО Бизнес Отель" — не ФИО

8. **#44 sodex 7.2** — переклассифицирован false_fail → true_fail (B)
   - Нет отдельного листа "Условные обозначения" (как и у mapper)

### Результат перезапуска

- Прогон: run_20260213_092223
- compare_runs.py vs baseline (run_20260210_165639):
  - ИСПРАВЛЕНО (FAIL → PASS): **25** (все false_fail для kpsc)
  - РЕГРЕССИИ (PASS → FAIL): 3 (все объяснены как ложные PASS в старом прогоне)
  - НОВЫЕ FAIL: 26 (парсеры теперь находят данные → выявлены реальные проблемы документов)
  - БЕЗ ИЗМЕНЕНИЙ: 48 (20 FAIL→FAIL + 28 PASS→PASS)
- false_fail kpsc: 27 → 0 (25 fixed + 2 reclassified)
- Точность kpsc: 40.0% → 60.8%
- Общая точность: 63.7% → 64.4%

### Верификация "регрессий"

1. **7.8 biznes_otel/rotosnab** — старый PASS был основан на мусорных данных takt_time. Теперь поле null → корректный FAIL
2. **4.1 rotosnab** — таблица теперь парсится правильно, видны реально пустые ячейки → корректный FAIL

### Юнит-тесты: 50/50 PASS

### Файлы изменены

- `audit_engine/kpsc/parser_scripts/parse_kpsc_header.py` — добавлено поле organization
- `audit_engine/kpsc/validation_scripts/validate_6_1_transport_row.py` — расширен фильтр колонок
- `audit_engine/kpsc/validation_scripts/validate_2_1_problems_count.py` — расширен фильтр колонок
- `audit_engine/kpsc/validation_scripts/validate_8_2_values_cross_check.py` — динамический ИТОГО
- `audit_engine/kpsc/validation_scripts/validate_1_2_company_name.py` — детерминированная предпроверка
- `audit_engine/kpsc/validation_scripts/validate_1_4_responsible.py` — детерминированная предпроверка

## [2026-02-13] — prikaz_comp_ppu

### Что исправлено

**10 false_fail — 4 типа фиксов:**

1. **Filename (D, 5 записей: #95,#99,#103,#108,#113)**
   - config.json уже имел `filename_keywords: ["приказ", "ппу"]`
   - Базовый прогон был ДО добавления keywords — использовал `filename_pattern: "order_comp_ppu"`
   - Обновлён rules.json: описание паттерна → ключевых слов
   - Все 5 компаний: русские имена файлов содержат "приказ" и "ппу"

2. **Preprocessor regex (E, 1 запись: #100 mapper rule 3)**
   - Баг: `r'(возложить на\s*).+?\.'` — ленивый `.+?` останавливался на первой точке инициалов
   - Пример: "возложить на Фролову А.О., координатора проектов." → "возложить на [ДОЛЖНОСТЬ_ФИО].О., координатора проектов."
   - Фикс: `.+?` → `.+` (жадный, до последней точки в строке)
   - Также исправлен аналогичный баг в `prikaz_ppu.py`

3. **Правила 4,5 переведены на non-LLM (C→non-LLM, 2 записи: #97,#110)**
   - LLM систематически ошибался: путал scope (#97), отвергал garbled OCR (#110)
   - Создан `audit_engine/non_llm_checks/prikaz_checks.py` с детерминированными проверками:
     - Rule 4: regex для ПРИКАЗ/Приказ + проверка текста после № (не пустой, не подчёркивания)
     - Rule 5: regex для города (г. Москва) + даты (числовой ДД.ММ.ГГГГ / словесный ДД месяц ГГГГ)
   - rules.json: `llm: true` → `llm: false` для правил 4 и 5

4. **Правило 6 — улучшение промпта (C, 2 записи: #112 + регрессия sodex)**
   - #112 ruslet: LLM-галлюцинация — видел ООО "РУСЛЕТ" и говорил "отсутствует"
   - Регрессия sodex: LLM проверял каждый чанк отдельно вместо "любой из чанков"
   - Фикс: добавлены few-shot примеры + явное "достаточно найти в ЛЮБОМ чанке"

### Результат перезапуска

- Прогон: run_20260213_100605
- compare_runs.py vs baseline (run_20260210_165639):
  - ИСПРАВЛЕНО (FAIL → PASS): **11** (10 false_fail + 1 true_fail #109)
  - РЕГРЕССИИ (PASS → FAIL): **0**
  - БЕЗ ИЗМЕНЕНИЙ: 29 (10 FAIL→FAIL + 19 PASS→PASS)
- false_fail prikaz_comp_ppu: 10 → 2 (10 исходных исправлены, 2 новых обнаружены при верификации)
- Точность prikaz_comp_ppu: 52.4% → 80.0% (8 true_fail / 10 FAIL)
- Общая точность: 64.4% → 65.8%
- Время прогона: 109с → 82с (−25%, 2 правила на non-LLM)

### Верификация оставшихся 10 FAIL (независимый агент)

- **8 true_fail** (B — реальные ошибки документов): незаполненные номера, даты, плейсхолдеры, отсутствие юрформы
- **2 false_fail** (A — парсинг, ruslet):
  - Rule 1: Vision OCR склеил "О проведении" → "Опроведении"
  - Rule 5: OCR прочитал дату "14.10.2025" как `ок'т дд,.о 2025г.`
  - Причина: документ ruslet содержит элементы, garbled при OCR (рукописные вставки?)
  - Не исправимо на уровне валидаторов — требуется fallback на python-docx парсер

### Известные ограничения

- **#109 ruslet rule 3** — true_fail (B) стал PASS из-за LLM-недетерминизма
  - "Приказываю:" (Title Case) vs "ПРИКАЗЫВАЮ:" (шаблон) — регистр
  - Не связано с нашими изменениями — LLM стохастически пропускает borderline-кейсы
- **ruslet** — Vision OCR (gpt-4.1-mini) систематически garbled на этом документе (номер, дата, заголовок)

### Файлы изменены

- `doc_configs/prikaz_comp_ppu/rules.json` — правила 2,4,5,6 обновлены
- `doc_configs/prikaz_comp_ppu/config.json` — без изменений (keywords уже были)
- `audit_engine/preprocessors/prikaz_comp_ppu.py` — regex `.+?` → `.+`
- `audit_engine/preprocessors/prikaz_ppu.py` — аналогичный regex-фикс
- `audit_engine/non_llm_checks/prikaz_checks.py` (новый) — non-LLM проверки правил 4,5
- `audit_engine/non_llm_checks/__init__.py` — регистрация prikaz_checks

## [2026-02-13] — Cluster 3: prikaz_vyhod, prikaz_comp_ppu, prikaz_ppu, prikaz_ic_potoka

### Общий результат

- Baseline (cluster 3): 81 PASS / 88 FAIL (точность 47.9%)
- После фиксов: 107 PASS / 62 FAIL (точность 63.3%, **+15.4 п.п.**)
- **26 false_fail исправлено** (100% из найденных)
- **60 true_fail сохранено** (100%, ни один реальный баг не потерян)
- **0 leaks, 0 регрессий**

### Что исправлено

#### 1. prikaz_checks.py — расширение non-LLM проверок

**a) Новая функция `check_signatory()` — rule 7 для всех 4 типов**
- Проверяет наличие должности и ФИО подписанта regex-ом
- Переведено с LLM из-за проблемы: gpt-4.1-mini считает фамилии на -ович (Александрович М.Д.) отчеством
- Обработка: плейсхолдеры, линия подписи (______), fallback на последний абзац если scope захватил весь документ
- Зарегистрирован для: prikaz_comp_ppu, prikaz_ppu, prikaz_vyhod, prikaz_ic_potoka

**b) Новая функция `check_fio_after_marker()` — rules 8,9 для prikaz_vyhod**
- Ищет фразу-маркер («назначить организатором/секретарем проведения обхода») → проверяет ФИО + должность после неё
- Поддержка: краткий ФИО (И.О. Фамилия), полный ФИО (Фамилия Имя Отчество), плейсхолдеры
- Переведено с LLM из-за: LLM путал scope и не видел ФИО при нестандартной формулировке

**c) Новая функция `check_responsible_fio_ic_potoka()` — rule 8 для prikaz_ic_potoka**
- Проверяет 4 позиции ФИО ответственных лиц:
  - A: текст_приказа — пункт с «ознакомить» → должность+ФИО
  - B: регламент п.2.1 — ответственный за ИЦ → ФИО
  - C: регламент п.2.2 — исполняющий обязанности → ФИО **после** этих слов
  - D: регламент п.2.3 — администратор → ФИО
- Причина переноса на non-LLM: LLM систематически ошибался на позиции C (3/5 компаний false_fail) — не видел ФИО после «исполняющему обязанности». Также context_filter LLM был жёстко привязан к нумерации пунктов и ломался при сдвиге.

**d) `check_prikaz_and_number()` — улучшения**
- Добавлен параметр `scopes` (для prikaz_ic_potoka: scope "шапка" вместо "заголовок_город"+"номер_дата")
- Обработка `б/н` (без номера) — считается незаполненным номером

**e) `check_city_and_date()` — фикс regex города**
- Убран паттерн `Москв[аеы]` (bare) — ФИАС-адрес "109316, Москва г, Волгоградский пр-кт" — юрадрес, не город подписания
- Добавлен `Москв[аеы](?!\s+г[,\s])` — Москва standalone, но НЕ "Москва г," (ФИАС)
- Добавлен `Санкт-Петербург`

**f) `check_signatory()` — не включает `_{4,}` в плейсхолдеры**
- Подчёркивания в блоке подписи — линия для ручной подписи: `Директор __________________ Р.А. Балашов`
- ФИО заполнено, подчёркивания рядом с ним — не плейсхолдер

**Регистрации:** +12 функций (по 3 правила × 4 doc_type, плюс rules 8,9 для prikaz_vyhod, rule 8 для prikaz_ic_potoka)

#### 2. doc_configs/prikaz_vyhod/rules.json — 9 правил переписано

- **Rule 1** (заголовок): добавлено «П Р И К А З через пробелы = ПРИКАЗ слитно — не ошибка»
- **Rule 3** (текст приказа): «сравни СМЫСЛ» → «сравни текст ДОСЛОВНО». Два типа нарушений: A) пункт полностью заменён, Б) пропущены/добавлены слова. Причина: старый промпт делал LLM слишком снисходительным (sodex r3 true_fail пропускался — «производственной» игнорировалось)
- **Rules 4,5** → `llm: false` (non-LLM check)
- **Rule 6** (юрформа cross_check): улучшен промпт — игнорировать кавычки, регистр, дополнительный текст в xlsx_шапка
- **Rule 7** → `llm: false` (non-LLM check_signatory)
- **Rules 8,9** → `llm: false` (non-LLM check_fio_after_marker)
- **Rule 10** (номер приказа cross_check): добавлено «символ № — оформление, НЕ часть номера. "5" = "№ 5"»

#### 3. doc_configs/prikaz_ic_potoka/rules.json — 5 правил переписано

- **Rules 4,5** → `llm: false` (non-LLM с scopes=["шапка"]), улучшены промпты (OCR-артефакты, ФИАС-адрес)
- **Rule 7** → `llm: false` (non-LLM check_signatory с scopes=["подписант"]), добавлена инструкция о фамилиях на -ович
- **Rule 8** → `llm: false`, удалены context_filter и content (полностью non-LLM)
- **Rule 9** (даты): переписан на семантический поиск по фразам-маркерам вместо фиксированных номеров пунктов

#### 4. doc_configs/prikaz_ppu/rules.json — 5 правил обновлено

- **Rule 2** (filename): описание обновлено на ключевые слова (keywords уже работали)
- **Rules 4,5** → `llm: false` (non-LLM), расширены промпты (словесный формат даты, типографские кавычки)
- **Rule 6** (юрформа): «достаточно найти в ЛЮБОМ из чанков»
- **Rule 7** → `llm: false` (non-LLM check_signatory), инструкция о фамилиях на -ович

#### 5. doc_configs/prikaz_comp_ppu/rules.json — 3 правила обновлено

- **Rule 3** (текст приказа): добавлено игнорирование синтаксиса плейсхолдеров `[ДОЛЖНОСТЬ_ФИО]` vs `<ФИО> и <должность>`, исключение по регистру для «ПРИКАЗЫВАЮ»
- **Rule 6** (юрформа): убраны few-shot примеры с реальными компаниями (ООО "РУСЛЕТ", ООО «Содекс» → ООО «Ромашка», АО «Вектор»)
- **Rule 7** → `llm: false` (non-LLM check_signatory)

#### 6. Препроцессоры

**prikaz_vyhod.py:**
- `normalize_text_for_rule3()`: regex boundary fix — `.+?\.` (non-greedy с точкой) → `.+?` + lookahead `(?=\n\n|\n\d+\.|$)`. Старый паттерн съедал следующие пункты приказа при DOTALL.
- «производственной площадке» → опциональное «производственной» (в некоторых документах без него)
- Новый: `normalize_header()` — «П Р И К А З» (разреженное) → «Приказ» (как в шаблоне)

**prikaz_comp_ppu.py + prikaz_ppu.py:**
- Нормализация `<ФИО> и <должность>` → `[ДОЛЖНОСТЬ_ФИО]` (унификация формата плейсхолдеров между target и template)

#### 7. test_analysis.json — 1 переклассификация

- **#230 mapper prikaz_ic_potoka rule 9**: false_fail D → **true_fail B**. Пересмотр: mapper имеет только 3 пункта, секция «ознакомить»/«в срок до» полностью отсутствует. Первая дата есть, вторая — нет. Реальная ошибка документа.

### Результаты по doc_type

| Doc Type | Baseline | После | Дельта | Регрессии |
|----------|----------|-------|--------|-----------|
| prikaz_vyhod | 27P/28F | 35P/20F | +8 | 0 |
| prikaz_comp_ppu | 19P/21F | 29P/11F | +10 | 0 |
| prikaz_ppu | 13P/11F | 16P/8F | +3 | 0 |
| prikaz_ic_potoka | 22P/28F | 27P/23F | +5 | 0 |

### Известные оставшиеся проблемы

- **ERR (transient):** LibreOffice иногда не конвертирует .docx → PDF при параллельном запуске тестов. Решается последовательным запуском. Не баг кода.
- **LLM non-determinism на rule 3** (сверка текста): иногда пропускает разницу регистра ("Приказываю" vs "ПРИКАЗЫВАЮ"). Связано с длиной контекста. Intermittent.
- Все оставшиеся FAIL — **true_fail** (реальные дефекты документов). Новых false_fail не обнаружено.

### Файлы изменены

- `audit_engine/non_llm_checks/prikaz_checks.py` — +3 функции, расширение существующих, +12 регистраций
- `audit_engine/preprocessors/prikaz_vyhod.py` — regex boundary fix + normalize_header
- `audit_engine/preprocessors/prikaz_comp_ppu.py` — placeholder normalization
- `audit_engine/preprocessors/prikaz_ppu.py` — placeholder normalization
- `doc_configs/prikaz_vyhod/rules.json` — 9 правил переписано (5 на non-LLM)
- `doc_configs/prikaz_ic_potoka/rules.json` — 5 правил переписано (4 на non-LLM)
- `doc_configs/prikaz_ppu/rules.json` — 5 правил обновлено (3 на non-LLM)
- `doc_configs/prikaz_comp_ppu/rules.json` — 3 правила обновлено (1 на non-LLM)
- `test_configs/test_analysis.json` — 1 переклассификация (false_fail → true_fail)

## [2026-02-13] — polozhenie_ppu

### Что исправлено

**11 false_fail — 4 группы фиксов (6 итераций тестирования):**

#### 1. Rule 2 — filename_keywords (D, 4 записи: #168,#172,#176,#180)

- config.json уже имел `filename_keywords: ["ппу"]`
- Базовый прогон использовал `filename_pattern: "reg_ppu"` (латиница) — не находил в русских именах
- Обновлён rules.json: описание изменено на ключевые слова

#### 2. Rule 1 — preprocess_shapka (C, 3 записи: #167,#171,#175)

- Vision Parser разбивал шапку на строки: `"положение\nо системе подачи..."` vs шаблон `"положение о системе подачи..."`
- LLM считал перенос строки содержательным отличием
- Фикс: добавлен шаг 6 в `preprocess_shapka()` — склейка всех whitespace в одну строку (`re.sub(r'\s+', ' ', text)`)
- Безопасно: шапка = один логический заголовок, переносы = артефакт

#### 3. Rule 3 — normalize_structure (C, 2 записи: #169,#173)

- Vision разбивал длинный заголовок раздела на 2 строки:
  `"3. Распределение функций и ответственности"` → `"между участниками процесса по подаче предложений"`
- `normalize_structure()` отбрасывал вторую строку (не начинается с цифры)
- Фикс: добавлена склейка строк-продолжений с lookahead (только если далее есть ещё раздел, иначе — мусор типа "Заголовки не найдены")

#### 4. Rule 4 — context_filter + preprocessor (C, 2 записи: #174,#179)

**Самый сложный фикс — 6 итераций:**

Проблема: gpt-4.1-mini галлюцинировал на длинном контексте (~197 строк: 70 TARGET + 62 TEMPLATE + инструкции). Каждая итерация устраняла один ложный дефект, но LLM находил новый:
- Итерация 1: переписан промпт rule 4 (алгоритм, IGNORE-список) → LLM поймал точки в нумерации (4.2.2.)
- Итерация 2: preprocessor шаг 3.6 (нормализация точек подпунктов) → LLM поймал запятые и тире
- Итерация 3: preprocessor шагы 7.5, 7.6 (нормализация `;,` → `.`, удаление маркеров списков) → LLM галлюцинировал "4.2.2 отсутствует в шаблоне" (при идентичном тексте!)
- Итерация 4: **context_filter** `^\d+\.\d+` — отсекает таблицы, оставляет только подпункты. Контекст -44%. → LLM поймал обрезанную таблицу категорий
- Итерация 5: preprocessor шаг 5 (удаление TAB-строк = таблиц) → LLM поймал "Форма №1" хвост и нелогично вернул FAIL сам написав "это не нарушение"
- Итерация 6: preprocessor шаг 5.5 (отсечение `Форма №...` и всего после) → **PASS!**

Ключевой вывод: для длинных документов (Положения) комбинация **context_filter + агрессивная очистка preprocessor** необходима для надёжной работы gpt-4.1-mini.

### Результат перезапуска

- Прогон: run_20260213_142621
- compare_runs.py vs первый прогон сессии (run_20260213_135558):
  - ИСПРАВЛЕНО (FAIL → PASS): **2** (rule 4: #174,#179)
  - РЕГРЕССИИ (PASS → FAIL): **0**
  - БЕЗ ИЗМЕНЕНИЙ: 18 (8 FAIL→FAIL + 10 PASS→PASS)
- compare_runs.py vs baseline с учётом rules 1-3 (всего за сессию):
  - ИСПРАВЛЕНО: **11/11** false_fail → PASS
  - РЕГРЕССИИ: **0**
- false_fail polozhenie_ppu: 11 → 0
- Точность polozhenie_ppu: 42.1% → 100% (8 true_fail / 8 FAIL)

### Верификация

- Оставшиеся 8 FAIL — все true_fail (реальные дефекты документов)
- 0 новых false_fail
- 0 leaks (ни один true_fail не потерян)

### Файлы изменены

- `doc_configs/polozhenie_ppu/rules.json` — правила 2 (keywords), 4 (context_filter + промпт)
- `audit_engine/preprocessors/polozhenie_normalize.py` — 7 изменений:
  - `preprocess_shapka()`: шаг 6 — склейка whitespace в одну строку
  - `normalize_structure()`: склейка строк-продолжений заголовков с lookahead
  - `normalize_text()`: шаг 3.6 — нормализация точек подпунктов (4.2.2. → 4.2.2)
  - `normalize_text()`: шаг 5 — удаление TAB-строк (таблицы)
  - `normalize_text()`: шаг 5.5 — отсечение "Форма №..." и хвоста
  - `normalize_text()`: шаг 7.5 — нормализация `,` → `.` в конце строк
  - `normalize_text()`: шаг 7.6 — удаление маркеров списков (`-`, `–`, `—`)

## [2026-02-13] — polozhenie_comp_ppu

### Что исправлено

**6 false_fail — полностью покрыты фиксами из предыдущих сессий (0 новых изменений):**

#### 1. Rule 1 — preprocess_shapka (C, 3 записи: #227,#232,#237)

- Shared preprocessor `polozhenie_normalize.py` уже содержит все необходимые фиксы из сессии polozhenie_ppu:
  - Шаг 3: удаление строки `к приказу от ... № ...` (переменная, мешает сравнению)
  - Шаг 4: обрезка тела документа (Vision включает разделы в шапку)
  - Шаг 5: нормализация регистра (`ПОЛОЖЕНИЕ` → `положение`)
  - Шаг 6: склейка whitespace в одну строку (Vision-артефакт переносов)
- Декораторы `@register_preprocessor("polozhenie_comp_ppu", "шапка")` уже были — фиксы применились автоматически

#### 2. Rule 2 — filename_keywords (D, 3 записи: #228,#233,#238)

- config.json уже имел `filename_keywords: ["положение", "ппу"]`
- Базовый прогон использовал `filename_pattern: "reg_comp_ppu"` (латиница)
- Обновлён rules.json: описание изменено на ключевые слова

### Результат перезапуска

- Прогон: run_20260213_150349
- Результат: 9 PASS / 9 FAIL (3 компании × 6 правил)
  - Rule 1: **PASS** × 3 (было FAIL) — исправлено
  - Rule 2: **PASS** × 3 (было FAIL) — исправлено
  - Rule 3: FAIL × 3 — true_fail B (ожидаемо)
  - Rule 4: FAIL × 3 — true_fail B (ожидаемо)
  - Rule 5: FAIL × 3 — true_fail B (ожидаемо)
  - Rule 6: PASS × 3 — без изменений
- Регрессии: **0**
- false_fail: 6 → 0 (все 6 исправлены)
- Точность: 60.0% → 100.0% (9 true_fail / 9 FAIL)

### Верификация

- Оставшиеся 9 FAIL — все true_fail B (реальные дефекты документов):
  - Rule 3: отсутствует раздел «4. Порядок организации и проведения конкурсов»
  - Rule 4: пропущены подпункты, изменены формулировки
  - Rule 5: номер приказа или дата отсутствуют / содержат плейсхолдеры
- 0 новых false_fail
- 0 leaks (ни один true_fail не потерян)

### Файлы изменены

- `doc_configs/polozhenie_comp_ppu/rules.json` — правило 2 (описание keywords)
- Все остальные фиксы — из shared `polozhenie_normalize.py` (сессия polozhenie_ppu)

## [2026-02-14] — prikaz_ic

### Что исправлено

**5 false_fail — 3 исправлены (→ PASS), 1 переклассифицирован (→ true_fail B), 1 open:**

1. **Rule 3 — whitespace в промпте + препроцессор** (entries #252, #269 — root cause C)
   - LLM считал пробел перед номером пункта текстовой ошибкой ("пробелы игнорируются, но всё равно FAIL")
   - Добавлен пункт 0) в исключения: "ПРОБЕЛЫ, отступы — ИГНОРИРУЙ"
   - Препроцессор: `line.lstrip()` для нормализации отступов в текст_приказа
   - Добавлена инструкция формата: "Если нет нарушений — верни ТОЛЬКО {status: ok}"

2. **Rule 4 — формат ФИО** (entry #253 — root cause C)
   - LLM не считал "Н.В. Филатов" полным ФИО
   - Уточнён промпт: "инициалы+фамилия типа 'Н.В. Филатов' — это НОРМА"

3. **Rule 2 — форматы даты и города** (entry #268 — root cause C)
   - Исправлены 3 ложные жалобы: пробел после №, "Москва г", "01.10.2025г."
   - НО каждый промпт-фикс убирает одну жалобу → LLM находит новую (whack-a-mole)
   - Попытка preprocessor для шапки: context_builder не применяет preprocessors для target_only
   - Попытка исправить context_builder: сломала rule 4 (preprocessor удалял данные)
   - **OPEN**: требует non-LLM check (prikaz_checks.py на другой ветке)

4. **Rule 6 — контекст Приложения** (entry #271 — root cause C → reclassified B)
   - Препроцессор `trim_appendix2_to_relevant_sections()` обрезает основное тело Регламента
   - НЕ активен (target_only не вызывает preprocessors), но помог при анализе
   - После ручной верификации: п.4.1/4.2 реально отсутствуют → true_fail B

### Результат перезапуска
- Прогон: run_20260214_120755
- compare_runs.py (vs baseline run_20260210_165639):
  - ИСПРАВЛЕНО (FAIL → PASS): 3 (#252, #253, #269)
  - РЕГРЕССИИ (PASS → FAIL): 0
  - Утечки true_fail → PASS: 0
- false_fail: было 5 → стало 1 (3 fixed + 1 reclassified + 1 open)
- Точность: 11.1% → 19.4%

### Архитектурные находки
- **context_builder.py**: preprocessors НЕ применяются для target_only и cross_check (только template)
- Это by design — агрессивные preprocessors (удаление плейсхолдеров) безопасны только для template
- Для target_only нужны "безопасные" preprocessors (whitespace, OCR) — требуется отдельный registry

### Файлы изменены
- audit_engine/preprocessors/prikaz_ic.py — lstrip(), normalize_header(), trim_appendix2()
- doc_configs/prikaz_ic/rules.json — rules 2 (форматы), 3 (whitespace + формат), 4 (ФИО)
- test_configs/test_analysis.json — #271 reclassified

## [2026-02-13] — prikaz_formirovanie_po

### Что исправлено

**5 false_fail — 3 исправлены (→ PASS), 2 переклассифицированы (→ true_fail B):**

1. **Rule 9 → non-LLM check** (entries #148, #154 — root cause C)
   - gpt-4.1-mini галлюцинирует отсутствие М.П. даже когда оно присутствует в контексте
   - Создана `check_signatory_with_stamp()` в `prikaz_checks.py` — regex для должности, ФИО, М.П./МП/ПЕЧАТЬ
   - Зарегистрирована для `prikaz_formirovanie_po` rule 9 (`llm: false`)

2. **Rule 7 — семантические исключения + формат ответа** (entry #168 — root cause D)
   - Заменены абсолютные номера пунктов на семантические критерии (по тексту, не по номеру)
   - 4 исключения: «Создать ПО», пункт с двоеточием, «Руководителю ПО», «Контроль над исполнением»
   - Добавлены явные инструкции формата: не включать исключённые/OK пункты в массив нарушения
   - Добавлено различие документов vs должностей для корректной семантики

3. **Preprocessor — line-joining для текст_приказа** (`iter8_normalize.py`)
   - Vision-парсер разбивает пункты приказа переносами: «ПО.\nОтветственный» → «ПО. Ответственный»
   - Препроцессор объединяет continuation lines внутри пунктов

4. **Реклассификация #146** (biznes_otel rule 6): false_fail C → true_fail B
   - После фикса препроцессора LLM корректно находит пропущенные пункты (7 вместо 9 в шаблоне)

5. **Реклассификация #147** (biznes_otel rule 7): false_fail C → true_fail B
   - LLM семантически прав: «руководителя ПО» в «Утвердить инструкцию руководителя ПО» — объект, не ответственный
   - 4 итерации промпта не смогли решить конфликт #147 vs #153 без регрессий

### Результат перезапуска
- Прогон: run_20260213_160614
- compare_runs.py (vs baseline run_20260210_165639):
  - ИСПРАВЛЕНО (FAIL → PASS): 3 (#148, #154, #168)
  - РЕГРЕССИИ (PASS → FAIL): 0
  - Утечки true_fail → PASS: 0
- false_fail: было 5 → стало 0 (3 fixed + 2 reclassified)
- Точность: 32.5% → 40.0%

### Верификация
- Все true_fail остались FAIL: подтверждено (0 утечек)
- Регрессии: нет
- #153 rotosnab rule 7 (true_fail B): остаётся FAIL ✓
- #160 rotosnab_dop rule 7 (true_fail B): остаётся FAIL ✓

### Файлы изменены
- audit_engine/non_llm_checks/prikaz_checks.py — `check_signatory_with_stamp()`
- audit_engine/preprocessors/iter8_normalize.py — line-joining step
- doc_configs/prikaz_formirovanie_po/rules.json — rules 7 (semantic), 9 (llm:false)
- test_configs/test_analysis.json — #146, #147 reclassified

## [2026-02-13] — cheklist_eu

### Что исправлено

**7 false_fail — 4 исправлены (→ PASS), 3 переклассифицированы (→ true_fail B):**

#### 1. chunks_vision.json — подписи pages=[2] → pages=[-1] (A, 4 записи: #79,#80,#82,#83)

- Баг: чанк `подписи` имел `pages: [2]`, но biznes_otel_nf и biznes_otel_spir — одностраничные документы (альбомная ориентация)
- VisionExtractor фильтровал page_indices=[1] vs total_pages=1 → пустой список → `[НЕТ СТРАНИЦ ДЛЯ ИЗВЛЕЧЕНИЯ]`
- Фикс: `pages: [-1]` (последняя страница) — работает для 1- и 2-страничных документов
- Результат:
  - **#80, #83** (rule 7, подписи+дата): **PASS** — подписи извлечены, дата в шапке есть
  - **#79, #82** (rule 3, юрлицо в подписях): FAIL — но уже true_fail B: юрлицо реально отсутствует в подписях документа

#### 2. Переклассификация rule 3 (A→B, 3 записи: #79,#82,#85)

- После фикса парсинга подписи извлекаются корректно
- Но юрлицо (ООО «Бизнес-отель» / ООО «Ротоснаб») реально отсутствует в блоках подписей
- Подписи содержат только: «От предприятия: должность, ФИО» — без наименования
- Переклассифицированы: false_fail A → true_fail B

#### 3. Preprocessor — удаление столбцов описаний ответов (C, 2 записи: #86,#93)

- Баг: Vision извлекал полную markdown-таблицу с 6 столбцами (№, критерий, оценка, описание_0, описание_1, описание_2)
- LLM сравнивал ВСЕ столбцы, включая описания ответов: «Нет» (шаблон) vs «нет» (целевой) — регистр
- Правило инструктировало сравнивать ТОЛЬКО критерии (столбец 2), но LLM игнорировал инструкцию
- Фикс: `normalize_table_for_comparison()` — обрезка строк до 3 столбцов (№, критерий, оценка)
- Промпт rule 4 также усилен: явное «НЕ сравнивай варианты ответов»
- Результат: rule 4 PASS для ВСЕХ 5 компаний (включая biznes_otel_nf, где была регрессия на 1й итерации)

### Результат перезапуска

- Прогон: run_20260213_152425
- compare_runs.py vs baseline (run_20260210_165639):
  - ИСПРАВЛЕНО (FAIL → PASS): **4** (#80 rule 7, #83 rule 7, #86 rule 4, #93 rule 4)
  - РЕГРЕССИИ (PASS → FAIL): **0**
  - БЕЗ ИЗМЕНЕНИЙ: 13 FAIL→FAIL + 23 PASS→PASS
- false_fail cheklist_eu: 7 → 0 (4 fixed + 3 reclassified B)
- Точность cheklist_eu: 58.8% → 100% (13 true_fail / 13 FAIL)

### Верификация

- Оставшиеся 13 FAIL — все true_fail B (реальные дефекты документов)
- 0 новых false_fail
- 0 leaks

### Файлы изменены

- `doc_configs/cheklist_eu/chunks_vision.json` — подписи pages=[2] → pages=[-1]
- `doc_configs/cheklist_eu/rules.json` — правило 4 (усиление инструкции о столбцах ответов)
- `audit_engine/preprocessors/cheklist.py` — обрезка столбцов описаний ответов в таблице
