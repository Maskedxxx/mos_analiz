# План рефакторинга: Vision Pipeline

**Дата:** 2026-01-30
**Версия:** 1.0

---

## 1. Обзор

### 1.1 Цели рефакторинга

1. Заменить хрупкие парсеры на базе python-docx универсальным Vision Pipeline
2. Добавить поддержку PDF и сканов
3. Унифицировать код между итерациями
4. Сохранить обратную совместимость с существующим LLM-аудитом

### 1.2 Затрагиваемые итерации

| Итерация | Документ | Текущий парсер | Чанки |
|----------|----------|----------------|-------|
| четвёртая_итерация_prikaz_ic | Приказ о создании ИЦ | parser_prikaz_ic_docs.py | 6 |
| пятая_итерация | Приказ о создании ИЦ (эл.) | parser_prikaz_ic_docs.py | 6 |
| шестая_итерация | Приказ о выходе + График | parser_prikaz_vyhod_docx.py + parser_grafik_xlsx.py | 5 |
| седьмая_итерация | Приказ о создании ИЦ потока | parser_prikaz_ic_potoka.py | 7 |

---

## 2. Архитектура

### 2.1 Структура модулей

```
мос_мониторинг/
├── shared/                           # НОВАЯ папка для общих модулей
│   ├── __init__.py
│   ├── vision_parser/
│   │   ├── __init__.py
│   │   ├── converter.py              # DOCX→PDF→PNG конвертация
│   │   ├── extractor.py              # Vision LLM извлечение
│   │   ├── aggregator.py             # Объединение чанков
│   │   └── config.py                 # Загрузка конфигурации
│   └── utils/
│       ├── __init__.py
│       ├── retry.py                  # Retry логика
│       └── logging.py                # Единый логгер
│
├── седьмая_итерация/
│   ├── config/
│   │   ├── tz_prikaz_ic_potoka.json  # Правила аудита (существует)
│   │   └── chunks_config.json        # НОВЫЙ: конфиг чанков для Vision
│   ├── scripts/
│   │   ├── parser_prikaz_ic_potoka.py  # ОСТАВИТЬ как legacy
│   │   ├── vision_parser_wrapper.py    # НОВЫЙ: обёртка Vision
│   │   └── run_prikaz_ic_llm_audit.py  # Модифицировать
│
└── ... (другие итерации аналогично)
```

### 2.2 Диаграмма зависимостей

```
                    ┌──────────────────┐
                    │  run_*_audit.py  │
                    └────────┬─────────┘
                             │
              ┌──────────────┼──────────────┐
              ▼              ▼              ▼
      ┌──────────────┐ ┌───────────┐ ┌────────────┐
      │ legacy       │ │  Vision   │ │  LLM       │
      │ parser_*.py  │ │  Parser   │ │  Audit     │
      │ (fallback)   │ │  (NEW)    │ │  Engine    │
      └──────────────┘ └─────┬─────┘ └────────────┘
                             │
              ┌──────────────┼──────────────┐
              ▼              ▼              ▼
      ┌──────────────┐ ┌───────────┐ ┌────────────┐
      │  converter   │ │ extractor │ │ aggregator │
      │  .py         │ │ .py       │ │ .py        │
      └──────────────┘ └───────────┘ └────────────┘
              │              │
              ▼              ▼
      ┌──────────────┐ ┌───────────┐
      │ LibreOffice  │ │ OpenAI    │
      │ poppler      │ │ API       │
      └──────────────┘ └───────────┘
```

---

## 3. Этапы внедрения

### Этап 1: Подготовка инфраструктуры (1-2 дня)

**Задачи:**

1. **Создать папку `shared/`**
   ```bash
   mkdir -p shared/vision_parser shared/utils
   touch shared/__init__.py
   touch shared/vision_parser/__init__.py
   touch shared/utils/__init__.py
   ```

2. **Создать базовые модули**

   **shared/vision_parser/converter.py:**
   - Функция `docx_to_pdf()` — вызов LibreOffice
   - Функция `pdf_to_images()` — вызов pdf2image
   - Класс `DocumentConverter`

   **shared/vision_parser/extractor.py:**
   - Функция `extract_text_from_image()` — вызов Vision API
   - Функция `process_pages_parallel()` — параллельная обработка
   - Класс `VisionExtractor`

   **shared/utils/retry.py:**
   - Декоратор `@retry_with_backoff`
   - Конфигурация retry для OpenAI API

3. **Обновить requirements.txt**
   ```
   # Добавить:
   pdf2image>=1.17.0
   tenacity>=8.0.0
   pytesseract>=0.3.10
   opencv-python>=4.8.0
   ```

4. **Тесты инфраструктуры**
   - Проверить установку LibreOffice
   - Проверить установку poppler
   - Тест конвертации DOCX → PDF → PNG

**Критерии завершения:**
- [ ] Все модули созданы
- [ ] Конвертация работает на тестовом документе
- [ ] Vision API вызывается успешно

---

### Этап 2: Пилот на седьмой_итерации (2-3 дня)

**Почему седьмая итерация:**
- Самая свежая итерация
- 7 чанков — хорошее покрытие функционала
- Есть готовые правила аудита

**Задачи:**

1. **Создать конфигурацию чанков**

   **седьмая_итерация/config/chunks_config.json:**
   ```json
   {
     "document_type": "prikaz_ic_potoka",
     "markers": {
       "шапка": null,
       "преамбула": "О создании информационного центра пилотного потока",
       "текст_приказа": "ПРИКАЗЫВАЮ",
       "приложение_1": "Приложение № 1",
       "приложение_2_регламент": "Приложение № 2",
       "лист_ознакомления": "ЛИСТ",
       "подписант": null
     },
     "required_chunks": [
       "шапка",
       "преамбула",
       "текст_приказа",
       "приложение_1",
       "приложение_2_регламент",
       "лист_ознакомления",
       "подписант"
     ]
   }
   ```

2. **Создать wrapper для Vision Parser**

   **седьмая_итерация/scripts/vision_parser_wrapper.py:**
   ```python
   import sys
   sys.path.insert(0, '/path/to/shared')

   from vision_parser import VisionParser

   def parse_prikaz_ic_potoka_vision(docx_path: str) -> dict:
       """
       Парсит документ через Vision Pipeline.
       Интерфейс совместим с parse_prikaz_ic_potoka().
       """
       parser = VisionParser(
           chunk_config_path="config/chunks_config.json"
       )
       return parser.parse(docx_path)
   ```

3. **Модифицировать run_prikaz_ic_llm_audit.py**

   ```python
   # Добавить флаг --use-vision
   parser.add_argument(
       "--use-vision",
       action="store_true",
       help="Использовать Vision Pipeline вместо python-docx"
   )

   # В main():
   if args.use_vision:
       from vision_parser_wrapper import parse_prikaz_ic_potoka_vision
       target_doc = parse_prikaz_ic_potoka_vision(args.target)
   else:
       from parser_prikaz_ic_potoka import parse_prikaz_ic_potoka
       target_doc = parse_prikaz_ic_potoka(args.target)
   ```

4. **A/B тестирование**

   ```bash
   # Тест с legacy парсером
   python run_prikaz_ic_llm_audit.py \
       --target documents/test.docx \
       --template templates/template.docx

   # Тест с Vision Pipeline
   python run_prikaz_ic_llm_audit.py \
       --target documents/test.docx \
       --template templates/template.docx \
       --use-vision

   # Сравнить результаты
   diff -u logs_result/session_legacy/ logs_result/session_vision/
   ```

**Критерии завершения:**
- [ ] Vision parser возвращает все 7 чанков
- [ ] Формат чанков идентичен legacy парсеру
- [ ] LLM-аудит проходит без ошибок
- [ ] Результаты аудита совпадают (или лучше)

---

### Этап 3: Миграция остальных итераций (3-4 дня)

**Порядок миграции:**

1. **пятая_итерация** (простая, 6 чанков, похожа на седьмую)
2. **четвёртая_итерация_prikaz_ic** (6 чанков)
3. **шестая_итерация** (сложнее: 2 документа, XLSX)

**Для каждой итерации:**

1. Создать `chunks_config.json`
2. Создать `vision_parser_wrapper.py`
3. Добавить флаг `--use-vision` в run_*_audit.py
4. Провести A/B тест
5. Задокументировать различия

**Специальные случаи:**

**шестая_итерация (XLSX График):**
- XLSX не конвертируется в изображения через LibreOffice
- Оставить openpyxl парсер для графика
- Vision только для DOCX (Приказ о выходе)

```python
# шестая_итерация/scripts/run_audit.py
def parse_documents(args):
    # DOCX через Vision
    prikaz_chunks = vision_parser.parse(args.prikaz_docx)

    # XLSX через openpyxl (как раньше)
    from parser_grafik_xlsx import parse_grafik
    grafik_chunks = parse_grafik(args.grafik_xlsx)

    return {**prikaz_chunks, **grafik_chunks}
```

**Критерии завершения:**
- [ ] Все 4 итерации работают с --use-vision
- [ ] Все A/B тесты пройдены
- [ ] Документация обновлена

---

### Этап 4: Стабилизация и документация (1-2 дня)

**Задачи:**

1. **Финальное тестирование**
   - Тест на 10+ документах каждого типа
   - Тест с повреждёнными документами
   - Тест retry логики при ошибках API
   - Нагрузочный тест (100 документов)

2. **Метрики и мониторинг**
   ```python
   # shared/utils/metrics.py
   class PipelineMetrics:
       conversion_time: float
       extraction_time: float
       total_cost: float
       pages_processed: int
       fallback_count: int
       errors: List[str]
   ```

3. **Документация**
   - README для shared/vision_parser/
   - Примеры использования
   - Troubleshooting guide

4. **Флаг --use-vision по умолчанию**
   ```python
   parser.add_argument(
       "--use-legacy",  # Инвертировать логику
       action="store_true",
       help="Использовать legacy python-docx парсер"
   )
   ```

**Критерии завершения:**
- [ ] 99% тестов проходят
- [ ] Документация готова
- [ ] Vision по умолчанию включён

---

## 4. Файлы для создания/изменения

### 4.1 Новые файлы

| Путь | Описание |
|------|----------|
| `shared/__init__.py` | Инициализация пакета |
| `shared/vision_parser/__init__.py` | Экспорт VisionParser |
| `shared/vision_parser/converter.py` | DOCX→PDF→PNG |
| `shared/vision_parser/extractor.py` | Vision LLM API |
| `shared/vision_parser/aggregator.py` | Merge чанков |
| `shared/vision_parser/fallback.py` | Tesseract OCR |
| `shared/vision_parser/config.py` | Загрузка конфигов |
| `shared/utils/retry.py` | Tenacity retry |
| `shared/utils/logging.py` | Единый логгер |
| `седьмая_итерация/config/chunks_config.json` | Конфиг чанков |
| `седьмая_итерация/scripts/vision_parser_wrapper.py` | Обёртка |
| ... (аналогично для других итераций) |

### 4.2 Изменяемые файлы

| Путь | Изменение |
|------|-----------|
| `requirements.txt` | Добавить новые зависимости |
| `седьмая_итерация/scripts/run_prikaz_ic_llm_audit.py` | Флаг --use-vision |
| `пятая_итерация/scripts/run_prikaz_ic_llm_audit.py` | Флаг --use-vision |
| `четвёртая_итерация_prikaz_ic/scripts/run_prikaz_ic_llm_audit.py` | Флаг --use-vision |
| `шестая_итерация/scripts/run_audit.py` | Флаг --use-vision |

### 4.3 Legacy файлы (сохранить)

Не удалять — оставить как fallback:
- `седьмая_итерация/scripts/parser_prikaz_ic_potoka.py`
- `пятая_итерация/scripts/parser_prikaz_ic_docs.py`
- `четвёртая_итерация_prikaz_ic/scripts/parser_prikaz_ic_docs.py`
- `шестая_итерация/scripts/parser_prikaz_vyhod_docx.py`

---

## 5. Риски и митигация

### 5.1 Технические риски

| Риск | Вероятность | Влияние | Митигация |
|------|-------------|---------|-----------|
| LibreOffice нестабилен | Средняя | Высокое | Retry + timeout |
| Vision API rate limits | Средняя | Среднее | Semaphore + backoff |
| Vision API недоступен | Низкая | Высокое | Retry 3 попытки + ошибка |
| Большие документы (>20 стр) | Низкая | Среднее | Пагинация, кэш |

### 5.2 Бизнес-риски

| Риск | Вероятность | Влияние | Митигация |
|------|-------------|---------|-----------|
| Увеличение стоимости | Низкая | Низкое | ~$0.002/документ |
| Увеличение времени | Средняя | Среднее | Параллельность |
| Регрессия качества | Низкая | Высокое | A/B тесты |

### 5.3 План отката

1. **Немедленный откат:**
   ```bash
   # Использовать legacy парсер
   python run_audit.py --use-legacy ...
   ```

2. **Полный откат:**
   - Удалить флаг --use-vision из run_audit.py
   - Вернуть import legacy парсера
   - Git revert коммита

---

## 6. Timeline

```
Неделя 1:
├── Пн-Вт: Этап 1 (инфраструктура)
├── Ср-Пт: Этап 2 (пилот седьмая_итерация)

Неделя 2:
├── Пн-Чт: Этап 3 (миграция остальных)
├── Пт: Этап 4 (стабилизация)

Итого: 7-10 рабочих дней
```

---

## 7. Checklist для каждой итерации

```markdown
## [название_итерации]

### Подготовка
- [ ] Создан chunks_config.json
- [ ] Создан vision_parser_wrapper.py
- [ ] Добавлен флаг --use-vision

### Тестирование
- [ ] Vision parser извлекает все чанки
- [ ] Формат совпадает с legacy
- [ ] LLM-аудит проходит
- [ ] A/B тест: результаты идентичны/лучше

### Документация
- [ ] README обновлён
- [ ] Примеры добавлены
```

---

## 8. Следующие шаги

1. **Начать с Этапа 1** — создание инфраструктуры
2. **Проверить окружение** — LibreOffice, poppler, Tesseract
3. **Создать первый конфиг чанков** для седьмой_итерации
4. **Провести пилотный запуск**
