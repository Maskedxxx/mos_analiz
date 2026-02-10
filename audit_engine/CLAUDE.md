# audit_engine — Ядро ИИ-аудита

Оркестратор проверки документов: Vision-парсинг → контекст → LLM/non-LLM проверки → отчёт.

## Поток данных

```
DOCX/PPTX → vision_parser/ (LibreOffice→PDF→Vision API) → Dict[чанк: текст]
    ↓
preprocessors/ (нормализация текста для сравнения с шаблоном)
    ↓
engine.py → параллельно по правилам из doc_configs/<тип>/rules.json:
  ├─ non_llm_checks/ — детерминированные проверки (Python)
  └─ llm_client.py  — LLM-проверки (OpenAI API)
    ↓
List[violations] → excel_reporter.py + logger.py
```

## Модули

| Файл | Назначение |
|------|-----------|
| `engine.py` | Оркестратор: загрузка конфигов, парсинг, запуск проверок (ThreadPoolExecutor) |
| `models.py` | Dataclass'ы: RuleSpec, AuditConfig, AuditResult |
| `llm_client.py` | Обёртка OpenAI API + парсинг JSON-ответов |
| `context_builder.py` | Формирование user_prompt по типу сравнения (target_only / template / cross_check) |
| `logger.py` | PipelineLogger — логирование промптов, ответов, артефактов в session_dir |
| `excel_reporter.py` | Экспорт violations → .xlsx |

## Поддиректории

| Папка | Что делает | Расширение |
|-------|-----------|------------|
| `vision_parser/` | DOCX→PDF→Vision API→JSON чанки | Редко (универсальный) |
| `preprocessors/` | Нормализация текста перед сравнением | `@register_preprocessor(doc_type, scope)` |
| `non_llm_checks/` | Проверки без LLM (имя файла, шапка, чек-лист) | `@register(doc_type, rule_index)` |
| `parsers/` | Парсеры вторичных файлов (XLSX) | `PARSERS` dict в `__init__.py` |
| `system_prompts/` | .txt промпты для LLM (default, cheklist, presentation) | Новый .txt + ссылка в config.json |
| `drivers/` | Спецмодуль: аудит драйверов производства (свой pipeline) | — |
| `kpsc/` | Спецмодуль: валидация КПСЦ (25+ валидаторов, свой pipeline) | — |

## Зависимости

- **openai** — Chat API + Vision API (AsyncOpenAI)
- **openpyxl**, **pandas** — Excel
- **pdf2image**, **Pillow** — PDF → PNG для Vision
- **LibreOffice** (headless) — DOCX → PDF конвертация
