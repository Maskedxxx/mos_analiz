# Система валидации документов КПСЦ

## Структура проекта

```
третья_итерация/
├── validation_rules.json          # Все правила валидации (25 правил)
├── validation_scripts/            # Отдельные валидаторы
│   ├── validate_1_1_kpsc_text.py
│   ├── validate_1_2_company_name.py
│   └── ... (всего 25 валидаторов)
├── parser_outputs/                # Результаты парсинга Excel
│   ├── kpsc_header_v2.json
│   ├── pokazateli_v3.json
│   └── ...
├── validation_outputs/            # Результаты валидации (JSON)
├── validation_logs/               # Детальные логи
├── run_all_validations.py         # 🔥 Мастер-скрипт
└── validation_report.xlsx         # 📊 Итоговый отчет
```

---

## Быстрый старт

### 1. Установка зависимостей

```bash
pip install pandas openpyxl openai
```

### 2. Установка API ключа

```bash
export OPENAI_API_KEY="ваш-ключ-здесь"
```

### 3. Запуск всех валидаторов

```bash
python run_all_validations.py
```

По умолчанию:
- Читает правила из `validation_rules.json`
- Использует данные из `parser_outputs/`
- Сохраняет результаты в `validation_outputs/`
- Создает отчет `validation_report.xlsx`
- Запускает 5 валидаторов параллельно

---

## Параметры запуска

### Базовые параметры

```bash
python run_all_validations.py \
  --parser-outputs parser_outputs \
  --output-dir validation_outputs \
  --report validation_report.xlsx \
  --parallel 5
```

### С детальным логированием

```bash
python run_all_validations.py --verbose
```

### Все параметры

| Параметр | Описание | По умолчанию |
|----------|----------|--------------|
| `--parser-outputs` | Папка с распарсенными данными | `parser_outputs` |
| `--output-dir` | Папка для результатов валидации | `validation_outputs` |
| `--report` | Файл итогового отчета | `validation_report.xlsx` |
| `--rules` | Файл с правилами | `validation_rules.json` |
| `--scripts-dir` | Папка с валидаторами | `validation_scripts` |
| `--parallel` | Количество параллельных процессов | `5` |
| `--verbose` | Детальное логирование | `False` |

---

## Запуск отдельного валидатора

Если нужно запустить один валидатор:

```bash
python validation_scripts/validate_1_1_kpsc_text.py \
  --parser-outputs parser_outputs \
  --output validation_outputs/validate_1_1.json \
  --verbose
```

---

## Формат итогового отчета

Excel-файл `validation_report.xlsx` содержит таблицу:

| rule_index | section | rule_title | status | discrepancy | duration_sec |
|------------|---------|------------|--------|-------------|--------------|
| 1.1 | РАЗДЕЛ 1 | Наличие слов 'КПСЦ ТС' | PASS | | 2.3 |
| 1.2 | РАЗДЕЛ 1 | Название предприятия | PASS | | 1.8 |
| 8.3 | РАЗДЕЛ 8 | Соответствие названий | FAIL | Отсутствуют: ВПП | 3.1 |

### Статусы

- **PASS** - Проверка пройдена успешно
- **FAIL** - Обнаружены несоответствия
- **ERROR** - Ошибка выполнения валидатора
- **TIMEOUT** - Превышен таймаут (5 минут)
- **MISSING** - Валидатор не найден

---

## Пример вывода

```
🚀 Запуск всех валидаторов...
   Правила: validation_rules.json
   Валидаторы: validation_scripts
   Параллельность: 5

📋 Загружено правил: 25

✅ [1/25] 1.1: PASS (2.3s)
✅ [2/25] 1.2: PASS (1.8s)
❌ [3/25] 8.3: FAIL (3.1s)
...

✅ Все валидаторы завершены за 45.2s

📊 Отчет сохранен: validation_report.xlsx

📈 Статистика:
   Всего правил: 25
   ✅ PASS: 20 (80.0%)
   ❌ FAIL: 5 (20.0%)

⏱  Общее время: 45.2s
```

---

## Структура правила валидации

Каждое правило в `validation_rules.json`:

```json
{
  "rule_index": "1.1",
  "rule_title": "Наличие слов 'КПСЦ ТС'",
  "section": "РАЗДЕЛ 1: ПРОВЕРКИ ЗАГОЛОВКА",
  "requirement_expert": "Имеются слова 'КПСЦ ТС'",
  "technical_description": "В заголовке должен присутствовать текст 'КПСЦ'",
  "source_files": ["kpsc_header_v2.json"],
  "validation_criteria": {
    "what_to_check": "...",
    "success_condition": "...",
    "error_condition": "..."
  }
}
```

---

## Разделы правил

1. **РАЗДЕЛ 1** (7 правил): Проверки заголовка и основных полей
2. **РАЗДЕЛ 2** (3 правила): Проверка блока "Проблемы"
3. **РАЗДЕЛ 4** (2 правила): Проверка блока "Расчет ВПП"
4. **РАЗДЕЛ 5** (1 правило): Проверка единиц измерения
5. **РАЗДЕЛ 6** (1 правило): Проверка строки "Транспортировки"
6. **РАЗДЕЛ 7** (8 правил): Проверка наличия листов Excel
7. **РАЗДЕЛ 8** (3 правила): Кросс-проверка между листами

**Всего: 25 правил**

---

## Troubleshooting

### Ошибка: "OPENAI_API_KEY not found"

```bash
export OPENAI_API_KEY="sk-..."
```

### Ошибка: "pandas not found"

```bash
pip install pandas openpyxl
```

### Валидатор не найден

Проверьте, что файл валидатора существует в `validation_scripts/`:

```bash
ls validation_scripts/validate_*.py
```

### Таймаут валидатора

Увеличьте таймаут в коде `run_all_validations.py` (строка 71):

```python
timeout=600  # 10 минут
```

---

## Логи

При использовании флага `--verbose`:

- Создается папка `validation_logs/`
- Для каждого правила создается лог-файл `validate_X_Y.log`
- Логи содержат:
  - Промпты для LLM
  - Ответы от LLM
  - Извлеченные данные
  - Таймстампы каждого шага

---

## Производительность

- **Последовательный запуск**: ~125s (25 валидаторов × 5s)
- **Параллельный запуск (5 потоков)**: ~45s
- **Параллельный запуск (10 потоков)**: ~30s

Рекомендуется: `--parallel 5` (оптимальный баланс скорости и нагрузки на API)
