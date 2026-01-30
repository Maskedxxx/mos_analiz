# План: Реализация LLM-валидаторов для проверок КПСЦ

## ✅ ТЕКУЩИЙ СТАТУС

**Что уже сделано:**
- ✅ **ЭТАП 1 ЗАВЕРШЕН:** Создан `validation_rules.json` со ВСЕМИ 23 правилами (разделы 1-7)
- ✅ **ЭТАП 2 ЗАВЕРШЕН:** Созданы ВСЕ 23 валидатора (проверен синтаксис)
- ✅ Протестирован подход с динамической подстановкой правил
- ✅ Валидатор 1.1 успешно работает и проходит тест

**Текущая архитектура:**
- Правила хранятся в `validation_rules.json` (23 правила, 333 строки)
- Каждый валидатор читает своё правило из JSON
- Промпт динамически формируется с требованиями из ТЗ
- Результат сохраняется в JSON формате
- Все 23 валидатора созданы и готовы к тестированию пользователем

**Следующий шаг:** Пользователь протестирует валидаторы с LLM вручную

---

## TODO LIST - Порядок выполнения задач

### ЭТАП 1: Заполнение validation_rules.json ✅ ЗАВЕРШЕН

**КРИТИЧНО:** Сначала заполнить все правила, потом писать скрипты!

- [x] **1.1 РАЗДЕЛ 1 (Заголовок):** Проверки 1.1-1.7 ✅
  - [x] 1.1 - Наличие "КПСЦ ТС" в заголовке ✅
  - [x] 1.2 - Название предприятия ООО "..." ✅
  - [x] 1.3 - Название потока ✅
  - [x] 1.4 - ФИО ответственного ✅
  - [x] 1.5 - Дата разработки ✅
  - [x] 1.6 - Период реализации ✅
  - [x] 1.7 - ФИО составителя ✅

- [x] **1.2 РАЗДЕЛ 2 (Проблемы):** Проверки 2.1-2.3 ✅
  - [x] 2.1 - Количество проблем (информационная) ✅
  - [x] 2.2 - Последовательность номеров ✅
  - [x] 2.3 - Наличие описания у каждой проблемы ✅

- [x] **1.3 РАЗДЕЛ 3 (Операции):** Проверка 3.1 ✅
  - [x] 3.1 - Сверка названий операций ✅

- [x] **1.4 РАЗДЕЛ 4 (ВПП):** Проверки 4.1-4.2 ✅
  - [x] 4.1 - Заполненность ячеек ВПП ✅
  - [x] 4.2 - Сумма ВПП = ИТОГО ✅

- [x] **1.5 РАЗДЕЛ 5 (Единицы):** Проверка 5.1 ✅
  - [x] 5.1 - Заполненность единиц измерения ✅

- [x] **1.6 РАЗДЕЛ 6 (Перемещения):** Проверка 6.1 ✅
  - [x] 6.1 - Наличие строки "Перемещения" ✅

- [x] **1.7 РАЗДЕЛ 7 (Листы):** Проверки 7.1-7.8 ✅
  - [x] 7.1 - Наличие листа "КПСЦ" ✅
  - [x] 7.2 - Наличие листа "Условные обозначения" ✅
  - [x] 7.3 - Наличие листа "Показатели" ✅
  - [x] 7.4 - Наличие листа "Оцифровка потерь" ✅
  - [x] 7.5 - Наличие листа "ПА-1" ✅
  - [x] 7.6 - Наличие листа "Диаграмма Спагетти" ✅
  - [x] 7.7 - Наличие листа "Перечень проблем по ДС" ✅
  - [x] 7.8 - Наличие листа "Расчет такта" ✅

**Итого:** 23 правила для заполнения - ВСЕ ЗАВЕРШЕНЫ ✅

---

### ЭТАП 2: Написание и тестирование валидаторов 🔧

**Порядок:** По приоритету от простых к сложным

#### 2.1 Простые валидаторы (15 скриптов)

**РАЗДЕЛ 1: Заголовок (7 скриптов)** ✅ ЗАВЕРШЕН
- [x] ✅ `validate_1_1_kpsc_text.py` - ГОТОВ, ПРОТЕСТИРОВАН
- [x] ✅ `validate_1_2_company_name.py` - ГОТОВ
- [x] ✅ `validate_1_3_flow_name.py` - ГОТОВ
- [x] ✅ `validate_1_4_responsible.py` - ГОТОВ
- [x] ✅ `validate_1_5_date_developed.py` - ГОТОВ
- [x] ✅ `validate_1_6_date_implementation.py` - ГОТОВ
- [x] ✅ `validate_1_7_compiled_by.py` - ГОТОВ

**РАЗДЕЛ 7: Листы (8 скриптов)** ✅ ЗАВЕРШЕН
- [x] ✅ `validate_7_1_sheet_kpsc.py` - ГОТОВ
- [x] ✅ `validate_7_2_sheet_legend.py` - ГОТОВ
- [x] ✅ `validate_7_3_sheet_pokazateli.py` - ГОТОВ
- [x] ✅ `validate_7_4_sheet_ocifrovka.py` - ГОТОВ
- [x] ✅ `validate_7_5_sheet_pa1.py` - ГОТОВ
- [x] ✅ `validate_7_6_sheet_spaghetti.py` - ГОТОВ
- [x] ✅ `validate_7_7_sheet_spaghetti_problems.py` - ГОТОВ
- [x] ✅ `validate_7_8_sheet_takt_time.py` - ГОТОВ

#### 2.2 Средние валидаторы (5 скриптов) ✅ ЗАВЕРШЕН

**РАЗДЕЛ 2: Проблемы (3 скрипта)** ✅
- [x] ✅ `validate_2_1_problems_count.py` - ГОТОВ
- [x] ✅ `validate_2_2_problems_sequence.py` - ГОТОВ
- [x] ✅ `validate_2_3_problems_description.py` - ГОТОВ

**РАЗДЕЛ 5 и 6 (2 скрипта)** ✅
- [x] ✅ `validate_5_1_units.py` - ГОТОВ
- [x] ✅ `validate_6_1_transport_row.py` - ГОТОВ

#### 2.3 Сложные валидаторы (3 скрипта) ✅ ЗАВЕРШЕН

**РАЗДЕЛ 3 и 4 (3 скрипта)** ✅
- [x] ✅ `validate_3_1_operations_names.py` - ГОТОВ
- [x] ✅ `validate_4_1_vpp_filled.py` - ГОТОВ
- [x] ✅ `validate_4_2_vpp_sum.py` - ГОТОВ

**Итого:** 23 валидатора для написания и тестирования - ВСЕ ЗАВЕРШЕНЫ ✅

---

### ЭТАП 3: Пакетный запуск и отчеты 📊

- [ ] **3.1** Создать `run_all_validations.sh` - скрипт для запуска всех валидаторов
- [ ] **3.2** Создать `generate_report.py` - агрегация результатов в единый отчет
- [ ] **3.3** Тестирование на реальных данных

---

## Контекст задачи

Реализовать систему валидации файла Excel "КПСЦ и Спагетти ТС_Предприятие.xlsx" с использованием LLM (GPT-4.1-mini) для проверки соответствия данных требованиям из ТЗ v2.

### Ключевые решения (согласованы с пользователем):
- **API ключ:** Переменная окружения `OPENAI_API_KEY`
- **Структура:** Каждый валидатор — самодостаточный скрипт (без общих модулей)
- **Расположение скриптов:** `/третья_итерация/validation_scripts/`
- **Результаты:** `/третья_итерация/validation_outputs/`
- **Температура LLM:** 0 (детерминированность)
- **Модель:** gpt-4o-mini

### Исходные данные:
- **ТЗ:** `ТЗ_Валидация_КПСЦ_v2_упрощенная.md` (23 проверки)
- **Парсеры:** `/третья_итерация/parser_scripts/`
- **JSON данные:** `/третья_итерация/parser_outputs/`

---

## Архитектура решения

### Принцип работы валидатора

Каждый валидатор — это Python-скрипт, который:

1. **Читает JSON** из parser_outputs (конкретные файлы зависят от проверки)
2. **Извлекает нужные данные** в читаемый формат
3. **Формирует промпт для LLM:**
   - Роль: "Ты эксперт по проверке документов КПСЦ"
   - Требование из ТЗ v2 (точная формулировка пункта проверки)
   - Фактические данные из JSON
   - Формат ответа (JSON schema)
4. **Вызывает OpenAI API** (gpt-4o-mini, temperature=0)
5. **Получает структурированный JSON-ответ:**
```json
{
  "rule_index": "1.1",
  "rule_title": "Наличие текста 'КПСЦ' в заголовке",
  "target_document": "КПСЦ и Спагетти ТС_Предприятие.xlsx",
  "status": "PASS" | "FAIL",
  "discrepancy": "Описание проблемы или пустая строка"
}
```
6. **Сохраняет результат** в validation_outputs/

---

## Структура проекта

```
третья_итерация/
├── parser_scripts/           # Существующие парсеры
├── parser_outputs/           # JSON результаты парсинга
├── validation_scripts/       # НОВАЯ ПАПКА: скрипты валидации
│   ├── validate_1_1_kpsc_text.py
│   ├── validate_1_2_company_name.py
│   ├── validate_1_3_flow_name.py
│   └── ... (всего 23 скрипта)
└── validation_outputs/        # НОВАЯ ПАПКА: результаты валидации
    ├── check_1_1_result.json
    ├── check_1_2_result.json
    └── ...
```

---

## Шаблон валидатора

### Базовая структура скрипта

Каждый валидатор следует единому паттерну:

```python
#!/usr/bin/env python3
"""
Валидатор 1.1: Проверка наличия текста "КПСЦ" в заголовке
Источник ТЗ: ТЗ_Валидация_КПСЦ_v2_упрощенная.md, Раздел 1, Проверка 1.1
"""
import os
import json
import argparse
from pathlib import Path
from openai import OpenAI

# Константы
RULE_INDEX = "1.1"
RULE_TITLE = "Наличие текста 'КПСЦ' в заголовке"
TARGET_DOC = "КПСЦ и Спагетти ТС_Предприятие.xlsx"

def load_data(parser_outputs_dir: Path) -> dict:
    """Загрузка необходимых JSON данных"""
    header_file = parser_outputs_dir / "kpsc_header_v2.json"
    with open(header_file, "r", encoding="utf-8") as f:
        return json.load(f)

def extract_relevant_data(data: dict) -> dict:
    """Извлечение данных для проверки"""
    return {
        "title": data["fields"].get("title", "")
    }

def build_prompt(extracted_data: dict) -> str:
    """Формирование промпта для LLM"""
    prompt = f"""
Ты эксперт по проверке документов КПСЦ (Карта Потока Создания Ценности).

ТРЕБОВАНИЕ ИЗ ТЗ v2:
Проверка 1.1: В заголовке (ячейка A1 или объединенная область A1:AE1)
должен присутствовать текст 'КПСЦ' (регистронезависимо).

ФАКТИЧЕСКИЕ ДАННЫЕ:
Заголовок документа: "{extracted_data['title']}"

ЗАДАНИЕ:
Проверь, содержится ли текст "КПСЦ" в заголовке (регистр не важен).

ФОРМАТ ОТВЕТА (строго JSON):
{{
  "rule_index": "{RULE_INDEX}",
  "rule_title": "{RULE_TITLE}",
  "target_document": "{TARGET_DOC}",
  "status": "PASS или FAIL",
  "discrepancy": "Описание проблемы если FAIL, иначе пустая строка"
}}

Верни только JSON, без дополнительного текста.
"""
    return prompt

def call_llm(prompt: str, api_key: str) -> dict:
    """Вызов OpenAI API"""
    client = OpenAI(api_key=api_key)

    response = client.chat.completions.create(
        model="gpt-4o-mini",
        messages=[
            {"role": "system", "content": "Ты эксперт по валидации документов КПСЦ. Отвечаешь строго в формате JSON."},
            {"role": "user", "content": prompt}
        ],
        temperature=0,
        response_format={"type": "json_object"}
    )

    result_text = response.choices[0].message.content
    return json.loads(result_text)

def save_result(result: dict, output_file: Path):
    """Сохранение результата"""
    output_file.parent.mkdir(parents=True, exist_ok=True)
    with open(output_file, "w", encoding="utf-8") as f:
        json.dump(result, f, ensure_ascii=False, indent=2)
    print(f"✓ Результат сохранен: {output_file}")

def main():
    parser = argparse.ArgumentParser(description=f"Валидатор {RULE_INDEX}: {RULE_TITLE}")
    parser.add_argument(
        "--parser-outputs",
        type=Path,
        required=True,
        help="Путь к папке с результатами парсинга (parser_outputs)"
    )
    parser.add_argument(
        "--output",
        type=Path,
        required=True,
        help="Путь для сохранения результата валидации (JSON файл)"
    )
    args = parser.parse_args()

    # Получение API ключа
    api_key = os.environ.get("OPENAI_API_KEY")
    if not api_key:
        raise ValueError("Не найден OPENAI_API_KEY в переменных окружения")

    # Шаги валидации
    print(f"[1/5] Загрузка данных из {args.parser_outputs}")
    data = load_data(args.parser_outputs)

    print(f"[2/5] Извлечение релевантных данных")
    extracted = extract_relevant_data(data)

    print(f"[3/5] Формирование промпта")
    prompt = build_prompt(extracted)

    print(f"[4/5] Вызов LLM (gpt-4o-mini)")
    result = call_llm(prompt, api_key)

    print(f"[5/5] Сохранение результата")
    save_result(result, args.output)

    # Вывод результата
    status_emoji = "✅" if result["status"] == "PASS" else "❌"
    print(f"\n{status_emoji} Проверка {RULE_INDEX}: {result['status']}")
    if result["discrepancy"]:
        print(f"   Проблема: {result['discrepancy']}")

if __name__ == "__main__":
    main()
```

---

## План реализации (пошагово)

### Этап 1: Создание первого валидатора (прототип)

**Цель:** Реализовать и протестировать один валидатор для проверки 1.1, чтобы отладить подход

**Шаги:**
1. Создать папку `третья_итерация/validation_scripts/`
2. Создать `validate_1_1_kpsc_text.py` (по шаблону выше)
3. Протестировать на реальных данных:
   ```bash
   export OPENAI_API_KEY="sk-..."
   python validation_scripts/validate_1_1_kpsc_text.py \
     --parser-outputs parser_outputs \
     --output validation_outputs/check_1_1_result.json
   ```
4. Проверить:
   - Корректность чтения JSON
   - Качество промпта LLM
   - Формат ответа
   - Точность проверки

**Ожидаемый результат:**
- Работающий скрипт
- JSON файл с результатом валидации
- Понимание того, что нужно скорректировать в шаблоне

---

### Этап 2: Создание остальных валидаторов

После успешного тестирования прототипа, создать оставшиеся 22 валидатора.

#### Группировка по сложности:

**Простые (аналогичны 1.1):**
- 1.1-1.7: Проверки заголовка (работа с kpsc_header_v2.json)
- 7.1-7.8: Наличие листов (проверка существования JSON файлов)

**Средние:**
- 2.1-2.3: Работа с таблицей проблем (ocifrovka_poteri_v2.json)
- 5.1: Единицы измерения (kpsc_table1_v2.json)
- 6.1: Наличие строки "Перемещения"

**Сложные:**
- 3.1: Сверка названий операций (несколько источников)
- 4.1-4.2: Расчеты ВПП (математика + LLM)

#### Приоритет создания:

1. **Сначала простые (1.1-1.7, 7.1-7.8)** - 15 валидаторов
   - Однотипная логика
   - Можно создать быстро

2. **Затем средние (2.1-2.3, 5.1, 6.1)** - 5 валидаторов

3. **В конце сложные (3.1, 4.1-4.2)** - 3 валидатора

---

## Особенности реализации отдельных проверок

### Проверка 1.1-1.7: Заголовок КПСЦ
**Источник:** `kpsc_header_v2.json` → `fields`
**Паттерн:** Простая проверка полей
**Промпт:** Требование + значение поля → PASS/FAIL

### Проверка 2.1: Количество проблем
**Источник:** `ocifrovka_poteri_v2.json` → `rows`
**Особенность:** Информационная проверка (всегда PASS)
**Промпт:** Подсчитать количество строк, вывести статистику

### Проверка 2.2-2.3: Нумерация проблем
**Источник:** `ocifrovka_poteri_v2.json` → `rows`
**Логика:**
- Извлечь номера из колонки B (col=2)
- Проверить последовательность
- Найти пропуски
**Промпт:** Список номеров → проверка последовательности

### Проверка 3.1: Сверка названий операций
**Источники:**
- `kpsc_table1_v2.json` → заголовки столбцов
- `kpsc_table1_v2.json` → строка "Наименование показателя"
**Логика:**
- Найти header row (обычно row=15)
- Найти строку с "Наименование показателя"
- Сравнить поколоночно
**Промпт:** Две таблицы → выявить несоответствия

### Проверка 4.1-4.2: ВПП
**Источник:** `kpsc_table1_v2.json` → строка "ВПП"
**Логика:**
- 4.1: Найти пустые ячейки
- 4.2: Проверить сумму = ИТОГО
**Промпт:** Таблица со значениями → проверка заполненности + математика

### Проверка 5.1: Единицы измерения
**Источник:** `kpsc_table1_v2.json` → колонка D
**Логика:** Проверить что все ячейки заполнены
**Промпт:** Список значений → найти пустые

### Проверка 6.1: Строка "Перемещения"
**Источник:** `kpsc_table1_v2.json` → rows
**Логика:** Найти строку с текстом "перемещени" (нечувствительно)
**Промпт:** Список строк → найти совпадение

### Проверки 7.1-7.8: Наличие листов
**Источники:** Различные JSON файлы
**Логика:**
- Проверить существование файла
- Проверить что есть данные (не пустой)
**Промпт:** Мета-информация о файле → подтвердить наличие

---

## Запуск и тестирование

### Запуск одной проверки:
```bash
export OPENAI_API_KEY="sk-..."

python третья_итерация/validation_scripts/validate_1_1_kpsc_text.py \
  --parser-outputs третья_итерация/parser_outputs \
  --output третья_итерация/validation_outputs/check_1_1_result.json
```

### Пакетный запуск (будет реализован позже):
```bash
# Создать скрипт run_all_validations.sh
for script in validation_scripts/validate_*.py; do
    python "$script" \
      --parser-outputs parser_outputs \
      --output "validation_outputs/$(basename $script .py)_result.json"
done
```

---

## Формат результата проверки

Каждый валидатор создает JSON файл с единым форматом:

```json
{
  "rule_index": "1.1",
  "rule_title": "Наличие текста 'КПСЦ' в заголовке",
  "target_document": "КПСЦ и Спагетти ТС_Предприятие.xlsx",
  "status": "PASS",
  "discrepancy": ""
}
```

или при ошибке:

```json
{
  "rule_index": "2.2",
  "rule_title": "Последовательность номеров проблем",
  "target_document": "КПСЦ и Спагетти ТС_Предприятие.xlsx",
  "status": "FAIL",
  "discrepancy": "В таблице 'Оцифровка потерь' нарушена последовательность. Пропущенные номера: [3, 8]"
}
```

---

## Преимущества LLM-подхода

1. **Гибкость:** LLM понимает вариации текста и контекст
2. **Толерантность к форматированию:** Игнорирует пробелы, регистр
3. **Человекочитаемые отчеты:** Понятные описания проблем
4. **Простота обновления:** Изменить требование = изменить промпт
5. **Сложные проверки:** LLM может анализировать соответствия, а не только наличие

---

## Следующие шаги после плана

1. **Реализация:** Создать validate_1_1_kpsc_text.py
2. **Тестирование:** Запустить на реальных данных
3. **Итерация:** Скорректировать промпт если нужно
4. **Масштабирование:** Создать остальные 22 валидатора по шаблону

---

## Критические файлы для реализации

### Для чтения (существующие):
- `ТЗ_Валидация_КПСЦ_v2_упрощенная.md` - требования
- `третья_итерация/parser_outputs/*.json` - данные
- `третья_итерация/parser_scripts/parse_kpsc_header.py` - пример парсера

### Для создания (новые):
- `третья_итерация/validation_scripts/validate_*.py` - 23 валидатора
- `третья_итерация/validation_outputs/*.json` - результаты

---

## 🤖 ИНСТРУКЦИИ ДЛЯ AI АГЕНТОВ

### Порядок работы (СТРОГО):

1. **СНАЧАЛА: Заполнить validation_rules.json**
   - Открыть `ТЗ_Валидация_КПСЦ_v2_упрощенная.md`
   - Для каждой проверки извлечь:
     - `requirement_expert` - требование эксперта
     - `technical_description` - техническое описание
     - `validation_criteria` - критерии проверки
     - `source_files` и `source_path` - источники данных
   - Добавить в `validation_rules.json`
   - Отметить в TODO LIST (раздел ЭТАП 1)

2. **ПОТОМ: Создать валидаторы**
   - Использовать `validate_1_1_kpsc_text.py` как шаблон
   - Изменить только:
     - `RULE_INDEX` константу
     - `RULE_TITLE` константу
     - `load_data()` - источники данных
     - `extract_relevant_data()` - извлекаемые поля
   - Промпт формируется автоматически из `validation_rules.json`

3. **ТЕСТИРОВАНИЕ каждого валидатора:**
   ```bash
   python validation_scripts/validate_X_Y_name.py \
     --parser-outputs parser_outputs \
     --output validation_outputs/check_X_Y_result.json
   ```
   - Проверить JSON результат
   - Если PASS - отметить ✅ в TODO LIST (раздел ЭТАП 2)
   - Если FAIL - проверить промпт и логику

4. **ТОЛЬКО ПОСЛЕ завершения всех валидаторов:**
   - Создать `run_all_validations.sh`
   - Создать `generate_report.py`

### Шаблон для копирования валидатора:

```python
#!/usr/bin/env python3
"""
Валидатор X.Y: [Название проверки]
Источник ТЗ: ТЗ_Валидация_КПСЦ_v2_упрощенная.md, Раздел X, Проверка X.Y
"""
import os
import json
import argparse
from pathlib import Path
from openai import OpenAI

# Константы - ИЗМЕНИТЬ для каждого валидатора
RULE_INDEX = "X.Y"
RULE_TITLE = "[Название проверки]"
TARGET_DOC = "КПСЦ и Спагетти ТС_Предприятие.xlsx"

def load_rule(rule_index: str) -> dict:
    """Загрузка правила проверки из JSON"""
    rules_file = Path(__file__).parent.parent / "validation_rules.json"
    with open(rules_file, "r", encoding="utf-8") as f:
        data = json.load(f)
    for rule in data["rules"]:
        if rule["rule_index"] == rule_index:
            return rule
    raise ValueError(f"Правило {rule_index} не найдено")

def load_data(parser_outputs_dir: Path) -> dict:
    """Загрузка необходимых JSON данных - ИЗМЕНИТЬ источники"""
    # Пример для заголовка:
    header_file = parser_outputs_dir / "kpsc_header_v2.json"
    with open(header_file, "r", encoding="utf-8") as f:
        return json.load(f)

def extract_relevant_data(data: dict) -> dict:
    """Извлечение данных для проверки - ИЗМЕНИТЬ поля"""
    # Пример:
    return {
        "title": data["fields"].get("title", "")
    }

def build_prompt(extracted_data: dict, rule: dict) -> str:
    """Формирование промпта - НЕ ИЗМЕНЯТЬ, работает автоматически"""
    prompt = f"""
Ты эксперт по проверке документов КПСЦ (Карта Потока Создания Ценности).

ТРЕБОВАНИЕ ЭКСПЕРТА:
{rule['requirement_expert']}

ТЕХНИЧЕСКОЕ ОПИСАНИЕ:
{rule['technical_description']}

КРИТЕРИИ ПРОВЕРКИ:
- Что проверять: {rule['validation_criteria']['what_to_check']}
- Условие успеха: {rule['validation_criteria']['success_condition']}
- Условие ошибки: {rule['validation_criteria']['error_condition']}

ФАКТИЧЕСКИЕ ДАННЫЕ:
{chr(10).join(f"{k}: {v}" for k, v in extracted_data.items())}

ЗАДАНИЕ:
Проверь соответствие фактических данных требованию эксперта.

ФОРМАТ ОТВЕТА (строго JSON):
{{
  "rule_index": "{rule['rule_index']}",
  "rule_title": "{rule['rule_title']}",
  "target_document": "{TARGET_DOC}",
  "status": "PASS или FAIL",
  "discrepancy": "Описание проблемы если FAIL, иначе пустая строка"
}}

Верни только JSON, без дополнительного текста.
"""
    return prompt

def call_llm(prompt: str, api_key: str) -> dict:
    """Вызов OpenAI API - НЕ ИЗМЕНЯТЬ"""
    client = OpenAI(api_key=api_key)
    response = client.chat.completions.create(
        model="gpt-4.1-mini-2025-04-14",
        messages=[
            {"role": "system", "content": "Ты эксперт по валидации документов КПСЦ. Отвечаешь строго в формате JSON."},
            {"role": "user", "content": prompt}
        ],
        temperature=0,
        response_format={"type": "json_object"}
    )
    return json.loads(response.choices[0].message.content)

def save_result(result: dict, output_file: Path):
    """Сохранение результата - НЕ ИЗМЕНЯТЬ"""
    output_file.parent.mkdir(parents=True, exist_ok=True)
    with open(output_file, "w", encoding="utf-8") as f:
        json.dump(result, f, ensure_ascii=False, indent=2)
    print(f"✓ Результат сохранен: {output_file}")

def main():
    parser = argparse.ArgumentParser(description=f"Валидатор {RULE_INDEX}: {RULE_TITLE}")
    parser.add_argument("--parser-outputs", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    args = parser.parse_args()

    api_key = os.environ.get("OPENAI_API_KEY")
    if not api_key:
        raise ValueError("Не найден OPENAI_API_KEY в переменных окружения")

    print(f"[1/6] Загрузка правила проверки {RULE_INDEX}")
    rule = load_rule(RULE_INDEX)

    print(f"[2/6] Загрузка данных из {args.parser_outputs}")
    data = load_data(args.parser_outputs)

    print("[3/6] Извлечение релевантных данных")
    extracted = extract_relevant_data(data)

    print("[4/6] Формирование промпта с правилом из ТЗ")
    prompt = build_prompt(extracted, rule)

    print("[5/6] Вызов LLM")
    result = call_llm(prompt, api_key)

    print("[6/6] Сохранение результата")
    save_result(result, args.output)

    status_emoji = "✅" if result["status"] == "PASS" else "❌"
    print(f"\n{status_emoji} Проверка {RULE_INDEX}: {result['status']}")
    if result["discrepancy"]:
        print(f"   Проблема: {result['discrepancy']}")

if __name__ == "__main__":
    main()
```

### Правила работы с TODO LIST:

- ✅ Отметка `[x]` означает задача ВЫПОЛНЕНА
- ⏳ Отметка `[ ]` означает задача В ОЖИДАНИИ
- Каждую выполненную задачу помечать галочкой
- После завершения раздела - обновить план

### Приоритеты:

1. **ВЫСОКИЙ:** ЭТАП 1 - Заполнение validation_rules.json (все 23 правила)
2. **СРЕДНИЙ:** ЭТАП 2.1 - Простые валидаторы (15 штук)
3. **СРЕДНИЙ:** ЭТАП 2.2 - Средние валидаторы (5 штук)
4. **НИЗКИЙ:** ЭТАП 2.3 - Сложные валидаторы (3 штуки)
5. **НИЗКИЙ:** ЭТАП 3 - Пакетный запуск

---

## Контрольные точки (Checkpoints)

✅ **Checkpoint 1:** validation_rules.json содержит все 23 правила - **ЗАВЕРШЕН 2025-12-03**
✅ **Checkpoint 2:** Все 23 валидатора созданы (синтаксис проверен) - **ЗАВЕРШЕН 2025-12-03**
⏳ **Checkpoint 3:** Все валидаторы протестированы с LLM пользователем
⏳ **Checkpoint 4:** Пакетный запуск всех 23 валидаторов успешен
⏳ **Checkpoint 5:** Итоговый отчет сгенерирован

