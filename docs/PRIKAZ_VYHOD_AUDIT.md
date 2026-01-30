# 6-я итерация: Приказ о проведении выхода + График обхода ОМ

## Обзор

LLM-аудит двух связанных документов:
- **DOCX**: Приказ о проведении выхода (3.6)
- **XLSX**: График обхода ОМ (приложение к приказу)

Особенность: проверка **cross-file** — сверка данных между DOCX и XLSX файлами.

## Документы

| Тип | Файл | Описание |
|-----|------|----------|
| Целевой DOCX | `documents/prikaz_vyhod_ok.docx` | Приказ о проведении выхода |
| Целевой XLSX | `documents/grafik_obhod_ok.xlsx` | График обхода производства |
| Шаблон DOCX | `templates/prikaz_vyhod_template.docx` | Шаблон приказа |

## Структура чанков

### DOCX (Приказ)

| Чанк | Содержимое |
|------|------------|
| `шапка` | Юр.форма + наименование (ООО «Пример») |
| `номер_дата` | № приказа + дата (№ 07/19 ... 22 июля 2025) |
| `заголовок_город` | Заголовок + город (Приказ + г. Москва) |
| `текст_приказа` | Пункты 1-4 с ПРИКАЗЫВАЮ |
| `должность_фио_подписанта` | Подпись (Генеральный директор ... ФИО) |

### XLSX (График)

| Чанк | Содержимое |
|------|------------|
| `шапка` | Наименование компании + "Приложение №1 к Приказу № XX" |
| `фио_должности` | Список сотрудников (ФИО — Должность) |

## Правила проверки (11 шт.)

| # | Правило | Тип | Источник |
|---|---------|-----|----------|
| 1 | Заголовок = шаблон | `template` | docx |
| 2 | Имя файла | `target_only` | docx (без LLM) |
| 3 | Текст приказа = шаблон | `template` | docx |
| 4 | Есть "ПРИКАЗ" + номер | `target_only` | docx |
| 5 | Есть город + дата | `target_only` | docx |
| 6 | Компания docx = xlsx | `cross_file` | docx + xlsx |
| 7 | Подписант заполнен | `target_only` | docx |
| 8 | Организатор (п.1) заполнен | `target_only` | docx |
| 9 | Секретарь (п.2) заполнен | `target_only` | docx |
| 10 | Номер приказа docx = xlsx | `cross_file` | docx + xlsx |
| 11 | ФИО в xlsx заполнены | `target_only` | xlsx |

### Типы сравнения

- **template** — сверка DOCX с шаблоном
- **target_only** — проверка только целевого файла
- **cross_file** — сверка данных между DOCX и XLSX

## Использование

```bash
cd шестая_итерация

# Базовый запуск
python start.py prikaz_vyhod_ok.docx grafik_obhod_ok.xlsx

# С указанием шаблона
python start.py prikaz_vyhod_ok.docx grafik_obhod_ok.xlsx prikaz_vyhod_template.docx

# Прямой вызов скрипта
python scripts/run_prikaz_vyhod_audit.py \
  --docx documents/prikaz_vyhod_ok.docx \
  --xlsx documents/grafik_obhod_ok.xlsx \
  --template templates/prikaz_vyhod_template.docx

# Отладка (показать промпты без вызова LLM)
python scripts/run_prikaz_vyhod_audit.py \
  --docx documents/prikaz_vyhod_ok.docx \
  --xlsx documents/grafik_obhod_ok.xlsx \
  --template templates/prikaz_vyhod_template.docx \
  --print-prompts
```

## Структура папки

```
шестая_итерация/
├── start.py                           # Точка входа
├── config/
│   └── tz_prikaz_vyhod.json           # 11 правил проверки
├── documents/
│   ├── prikaz_vyhod_ok.docx           # Целевой DOCX
│   └── grafik_obhod_ok.xlsx           # Целевой XLSX
├── templates/
│   ├── prikaz_vyhod_template.docx     # Шаблон DOCX
│   └── grafik_obhod_template.xlsx     # (не используется)
├── scripts/
│   ├── parser_prikaz_vyhod_docx.py    # Парсер DOCX
│   ├── parser_grafik_xlsx.py          # Парсер XLSX
│   └── run_prikaz_vyhod_audit.py      # Основной аудит
└── logs_result/
    └── session_YYYYMMDD_HHMMSS/       # Результаты сессии
        ├── audit_result.xlsx
        ├── final_results.json
        ├── rules_summary.json
        ├── pipeline.log
        ├── parsed_docs/
        ├── prompts/
        └── responses/
```

## Логирование сессии

Каждый запуск создаёт папку `logs_result/session_YYYYMMDD_HHMMSS/`:

| Файл | Описание |
|------|----------|
| `audit_result.xlsx` | Excel отчёт с нарушениями |
| `final_results.json` | JSON массив нарушений |
| `rules_summary.json` | Сводка: статус каждого правила |
| `pipeline.log` | Лог выполнения с таймстампами |
| `parsed_docs/*.json` | Распарсенные документы |
| `prompts/rule_XX_prompt.txt` | Промпты для LLM |
| `responses/rule_XX_response.txt` | Ответы LLM |

## Препроцессинг

Для правила #3 (сверка текста с шаблоном) применяется нормализация плейсхолдеров:

```python
def normalize_text_for_rule3(text):
    # п.1: должность + ФИО организатора → [ДОЛЖНОСТЬ_ФИО]
    # п.2: должность + ФИО секретаря → [ДОЛЖНОСТЬ_ФИО]
```

Это позволяет сравнивать заполненный документ с шаблоном без ложных срабатываний.

## API

- **Модель**: gpt-4.1-mini (OpenAI)
- **Температура**: 0.0
- **Параллельность**: 4 потока
