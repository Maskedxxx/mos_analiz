# doc_configs — Конфигурации типов документов

16 типов документов. Каждый тип — отдельная папка с полным набором правил. Новый тип добавляется без изменения кода ядра.

## Структура типа

```
doc_configs/<doc_type>/
├── config.json           # Метаданные: модель, промпт, max_workers, температура
├── rules.json            # Правила проверки (index, scope, compare, llm, content)
├── chunks_vision.json    # Разметка для Vision-парсера (чанки, страницы, промпты)
└── template/
    ├── template.docx     # Эталонный шаблон
    └── template_cached.json  # Кэш Vision-парсинга (авто, по SHA256)
```

## config.json — ключевые поля

| Поле | Тип | Описание |
|------|-----|----------|
| `doc_type` | str | ID типа = имя папки |
| `doc_title` | str | Человекочитаемое название |
| `model` | str | Модель OpenAI (обычно `gpt-4.1-mini`) |
| `system_prompt` | str | Файл из `audit_engine/system_prompts/` |
| `max_workers` | int | Параллельные потоки LLM (1–5) |
| `temperature` | float | Обычно 0.0 |
| `engine` | str? | `kpsc` или `drivers` для спецмодулей, иначе Vision pipeline |

## rules.json — типы проверок

| compare | Что сравнивает | Когда использовать |
|---------|---------------|-------------------|
| `target_only` | Только целевой документ | Проверка реквизитов, наличия элементов |
| `template` | Целевой ↔ шаблон | Соответствие структуре, формулировкам |
| `cross_check` | Чанки целевого между собой | Сверка шапки с подписями, дат в разных местах |

Поле `llm: false` → non-LLM проверка (Python), `llm: true` → LLM-проверка.

## Добавление нового типа

1. Создать папку `doc_configs/<новый_тип>/`
2. Заполнить `config.json`, `rules.json`, `chunks_vision.json`
3. Положить эталон в `template/template.docx`
4. Проверить: `python run_audit.py --list-types`
5. Запустить: `python run_audit.py --doc-type <новый_тип> --target <файл>`
