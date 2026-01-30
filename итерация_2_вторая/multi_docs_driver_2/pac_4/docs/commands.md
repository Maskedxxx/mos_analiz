# Команды запуска скриптов

## Проверка приказов

### Проверка приказа о конкурсах ППУ

```bash
python scripts/run_order_llm_audit_v2.py \
  --config config/parser_paths.json \
  --target-id order_version \
  --template-id order_template \
  --tz-id order_tz \
  --out-xlsx results/order_llm_audit.xlsx
```

### Проверка приказа о ППУ

```bash
python scripts/run_order_llm_audit_v2.py \
  --config config/parser_paths.json \
  --target-id order_ppu_version \
  --template-id order_ppu_template \
  --tz-id order_ppu_tz \
  --out-xlsx results/order_ppu_llm_audit.xlsx
```

### Параметры

- `--config` — путь к конфигурации (по умолчанию: `config/parser_paths.json`)
- `--target-id` — ID целевого документа в конфиге
- `--template-id` — ID шаблона в конфиге
- `--tz-id` — ID ТЗ в конфиге
- `--model` — модель OpenAI (по умолчанию: `gpt-4.1-2025-04-14`)
- `--temperature` — температура генерации (по умолчанию: `0.0`)
- `--out-xlsx` — путь для сохранения результата
- `--print-prompt` — показать промпт для первого правила и выйти
- `--print-prompts-all` — показать промпты для всех правил и выйти

---

## Проверка регламентов

### Проверка регламента о конкурсах ППУ

```bash
python scripts/run_reg_comp_llm_audit.py \
  --config config/parser_paths.json \
  --target-id reg_comp_ppu_version \
  --template-id reg_comp_ppu_template \
  --tz-id reg_comp_ppu_tz \
  --out-xlsx results/reg_comp_ppu_llm_audit.xlsx
```

### Проверка регламента о ППУ

```bash
python scripts/run_reg_comp_llm_audit.py \
  --config config/parser_paths.json \
  --target-id reg_ppu_version \
  --template-id reg_ppu_template \
  --tz-id reg_ppu_tz \
  --out-xlsx results/reg_ppu_llm_audit.xlsx
```

### Параметры

- `--config` — путь к конфигурации
- `--target-id` — ID целевого документа
- `--template-id` — ID шаблона
- `--tz-id` — ID ТЗ
- `--model` — модель OpenAI
- `--temperature` — температура генерации
- `--out-xlsx` — путь для сохранения результата
- `--print-prompt` — показать промпт для первого правила
- `--print-prompts-all` — показать промпты для всех правил

---

## Пакетная проверка

Проверяет соответствие между 4 документами (номера/даты приказов и должности/ФИО).

```bash
python scripts/run_package_llm_audit.py \
  --config config/parser_paths.json \
  --out-xlsx results/package_llm_audit.xlsx
```

### Параметры

- `--config` — путь к конфигурации
- `--order-comp-id` — ID приказа о конкурсах ППУ (по умолчанию: `order_version`)
- `--reg-comp-id` — ID регламента о конкурсах ППУ (по умолчанию: `reg_comp_ppu_version`)
- `--order-id` — ID приказа о ППУ (по умолчанию: `order_ppu_version`)
- `--reg-id` — ID регламента о ППУ (по умолчанию: `reg_ppu_version`)
- `--model` — модель OpenAI
- `--temperature` — температура генерации (по умолчанию: `0.1`)
- `--out-xlsx` — путь для сохранения результата
- `--print-prompts` — показать промпты и выйти
- `--debug` — включить отладочные логи

---

## Полная проверка (все типы)

Запускает все проверки последовательно с сохранением результатов в уникальную папку.

```bash
python scripts/run_all_llm_audit.py \
  --config config/parser_paths.json \
  --runs-dir results/runs \
  --max-workers 4
```

### Параметры

- `--config` — путь к конфигурации
- `--model` — модель OpenAI (по умолчанию: `gpt-4.1-2025-04-14`)
- `--temperature` — температура генерации
- `--runs-dir` — базовая папка для результатов (по умолчанию: `results/runs`)
- `--pipeline-id` — указать ID пайплайна (иначе сгенерируется автоматически)
- `--max-workers` — максимум параллельных процессов (по умолчанию: `4`)
- `--print-prompts` — только показать промпты для всех одиночных запусков и выйти

### Результат

Создается папка с уникальным ID (например, `results/runs/20251126_143052_a3f2b1/`) с файлами:

- `order_llm_audit.xlsx` — проверка приказа о конкурсах ППУ
- `order_ppu_llm_audit.xlsx` — проверка приказа о ППУ
- `reg_comp_ppu_llm_audit.xlsx` — проверка регламента о конкурсах ППУ
- `reg_ppu_llm_audit.xlsx` — проверка регламента о ППУ
- `package_llm_audit.xlsx` — пакетная проверка всех 4 документов

---

## Настройка окружения

### Установка зависимостей

```bash
pip install openai pandas python-docx openpyxl
```

### Установка API ключа

```bash
export OPENAI_API_KEY="your-api-key-here"
```

---

## Примеры использования

### Быстрая проверка с показом промптов

```bash
python scripts/run_order_llm_audit_v2.py --print-prompt
```

### Проверка с кастомной моделью

```bash
python scripts/run_all_llm_audit.py --model gpt-4o --temperature 0.2
```

### Полная проверка с отладкой

```bash
python scripts/run_package_llm_audit.py --debug
```
