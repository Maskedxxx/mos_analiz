# Changelog v2 — Регрессионный прогон и фиксы

## Прогон

- **Baseline:** run_20260210_165639 (289 FAIL, 184 PASS, 0 ERR)
- **Регрессионный:** run_20260214_134634 (226 FAIL, 292 PASS, 0 ERR)
- **Файлов:** 51, **Время:** ~42 мин
- **Ветка фиксов:** `fix/regression-fixes` (от `main`)

## Результат compare_runs.py

- FAIL → PASS: 95 (93 ожидаемых false_fail + 2 маскировки true_fail)
- Регрессии PASS → FAIL: 6 (3 наш код + 3 LLM/Vision стохастика)
- Новые FAIL: 26 (все kpsc, NOT_RUN→FAIL — улучшенные парсеры)
- Неисправленные false_fail: 3
- Улучшение: 289→226 FAIL (+63 меньше)

---

## Расследование (завершено)

### Категория 1: Регрессии PASS→FAIL

| # | Doc/Rule | Причина | Классификация | Статус |
|---|----------|---------|---------------|--------|
| 1 | kpsc/biznes_otel 7.8 | Рефакторинг парсера удалил поле `takt_time` из `extract_fields()` | OUR_FIX | **TODO** |
| 2 | kpsc/rotosnab 4.1 | `sheet_finder` берёт "КПСЦ ТС" (20 колонок) вместо "КПСЦ" (17), validator не фильтрует null-колонки | OUR_FIX | **TODO** |
| 3 | kpsc/rotosnab 7.8 | Та же причина что #1 — `takt_time` удалено из парсера | OUR_FIX | **TODO** |
| 4 | polozhenie_po/biznes_otel 10 | LLM стохастика на рукописном OCR-мусоре | FLAKY | Пропуск |
| 5 | prikaz_ic/ruslet 5 | LLM цепляется за формат примера "№15БП" | FLAKY | Пропуск |
| 6 | prikaz_ic/ruslet 8 | Vision выдал "Периодичность актуализации документа" | FLAKY | Пропуск |

**Фикс #1,3 (takt_time):**
- Файл: `audit_engine/kpsc/parser_scripts/parse_kpsc_header.py`
- Проблема: `extract_fields()` переписана без `takt_time`
- Валидатор: `audit_engine/kpsc/validators/validate_7_8_sheet_takt_time.py` (строка 52) ищет `data["data"]["fields"]["takt_time"]`
- Решение: вернуть `takt_time` с динамическим поиском по лейблу ИЛИ переписать валидатор 7.8

**Фикс #2 (validator 4.1):**
- Файл: `audit_engine/kpsc/validators/validate_4_1_vpp_filled.py` (строки 96-113)
- Проблема: итерирует все колонки >= 5, не фильтрует null-заголовки
- Решение: добавить `if operation_name is None: continue` (пропускать колонки без заголовка)

### Категория 2: Маскировки true_fail→PASS

| Entry | Doc/Rule | Причина | Статус |
|-------|----------|---------|--------|
| #166 | prikaz_formirovanie_po/ruslet 5 | Vision галлюцинировал "г. Москва" (города нет в DOCX). Промпт НЕ менялся. | FLAKY / future: non-LLM |
| #122 | prikaz_ic_potoka/mapper 3 | **БАГ тест-раннера**: LLM timeout → `_parse_vision_results()` ставит PASS вместо ERROR | **TODO** — фикс тест-раннера |

**Фикс #122 (баг тест-раннера):**
- Файл: `run_tests.py`, строки 268-298 (`_parse_vision_results()`)
- Проблема: функция ставит PASS по умолчанию, ERROR/timeout не проверяются
- Файлы ошибок: `session/responses/rule_NN_error.txt` — не читаются
- Решение: после установки PASS по умолчанию, проверить `responses/rule_*_error.txt` → если есть, ставить ERROR

### Категория 3: Неисправленные false_fail

| Entry | Doc/Rule | Причина | Статус |
|-------|----------|---------|--------|
| #192 | polozhenie_po/biznes_otel 5 | Vision API деградация (garbled output) | FLAKY |
| #268 | prikaz_ic/ruslet 2 | Фикс частично сработал (1/2), LLM нашёл новую проблему с ФИАС-адресом | **TODO** |
| #135 | prikaz_ic_potoka/ruslet 8 | Non-LLM нашёл true_fail на позиции A (нет ФИО в п.4 "ознакомить") | **TODO** → переклассификация |

**Фикс #268 (prikaz_ic rule 2 промпт):**
- Файл: `doc_configs/prikaz_ic/rules.json`, rule 2, пункт 4 (город)
- Проблема: LLM видит ФИАС-адрес "109316, Москва г, ..." и считает город "частью адреса"
- Решение: добавить "Полный почтовый/юридический адрес с индексом, содержащий название города — ВАЛИДНОЕ указание города"

**Переклассификация #135:**
- Entry #135 в test_analysis.json: verdict false_fail → true_fail
- Причина: non-LLM корректно обнаружил реальное нарушение (п.4 без ФИО)

### Категория 4: Новые FAIL (26 kpsc)

Все 26 — NOT_RUN→FAIL. Парсеры в baseline падали, после рефакторинга работают. Все нарушения — реальные (true_fail).
- rule 1.x (header fields): 13 штук — поля None в mapper/sodex/ruslet
- rule 2.2 (нумерация проблем): 5 штук — все компании
- rule 5.1 (пустые ячейки): 2 штуки
- rule 7.8 (takt_time): 1 штука
- rule 8.x (кросс-валидация): 3 штуки

Все нужно записать как new true_fail через `submit_verdict`, но это не блокирует фиксы.

---

## План фиксов (батчи)

### Батч 1: kpsc регрессии (приоритет высокий)
- [ ] Фикс takt_time в parse_kpsc_header.py (регрессии 7.8 для biznes_otel + rotosnab + ruslet)
- [ ] Фикс validator 4.1 — фильтр null-колонок (регрессия rotosnab 4.1)
- [ ] Точечный тест: `run_tests.py --doc-type kpsc`
- [ ] Сравнение: `compare_runs.py --old run_20260210_165639 --new <run> --doc-type kpsc`

### Батч 2: баг тест-раннера (приоритет высокий)
- [ ] Фикс `_parse_vision_results()` в run_tests.py — обработка ERROR
- [ ] Нет точечного теста (инфраструктурный код)

### Батч 3: промпт prikaz_ic rule 2 (приоритет средний)
- [ ] Обновить правило 2 пункт 4 в `doc_configs/prikaz_ic/rules.json`
- [ ] Точечный тест: `run_tests.py --doc-type prikaz_ic --company ruslet`
- [ ] Сравнение: `compare_runs.py --old run_20260210_165639 --new <run> --doc-type prikaz_ic`

### Батч 4: переклассификация + запись новых true_fail
- [ ] Entry #135: submit_verdict → true_fail
- [ ] 26 новых kpsc FAIL: submit_verdict → true_fail (batch)

---

## FLAKY-тесты (не чиним)

| Doc/Rule | Причина | Тип |
|----------|---------|-----|
| polozhenie_po/biznes_otel rule 10 | Рукописный OCR-мусор, LLM flip-flop | LLM |
| prikaz_ic/ruslet rule 5 | LLM цепляется за формат примера | LLM |
| prikaz_ic/ruslet rule 8 | Vision OCR вариация в названии столбца | Vision |
| prikaz_formirovanie_po/ruslet rule 5 | Vision галлюцинирует город | Vision |
| polozhenie_po/biznes_otel rule 5 | Vision garbled output при повторном OCR | Vision |
