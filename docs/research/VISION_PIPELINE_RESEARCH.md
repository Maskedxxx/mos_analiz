# Universal Vision Pipeline: Исследование

**Дата:** 2026-01-30
**Версия:** 1.0
**Статус:** Исследование завершено

---

## 1. Обзор архитектуры

### 1.1 Текущая проблема

Существующие парсеры Word документов (`parser_prikaz_ic_docs.py`, `parser_prikaz_vyhod_docx.py`, `parser_prikaz_ic_potoka.py`) используют библиотеку `python-docx` для извлечения текста. Это работает, но имеет ограничения:

- Не работает со сканами и PDF
- Сложная логика поиска маркеров (fuzzy matching)
- Таблицы извлекаются упрощённо
- Каждый тип документа требует отдельного парсера

### 1.2 Предлагаемая архитектура

```
┌─────────────────────────────────────────────────────────────────┐
│                    Universal Vision Pipeline                     │
├─────────────────────────────────────────────────────────────────┤
│                                                                 │
│  ВХОД: DOCX / PDF / скан (PNG, JPEG)                           │
│         │                                                       │
│         ▼                                                       │
│  ┌─────────────────────────────────────────┐                   │
│  │  КОНВЕРТАЦИЯ                             │                   │
│  │  • DOCX → PDF (LibreOffice headless)    │                   │
│  │  • PDF → PNG[] (pdf2image + poppler)    │                   │
│  │  • DPI: 200-300 для оптимального OCR    │                   │
│  └─────────────────────────────────────────┘                   │
│         │                                                       │
│         ▼                                                       │
│  ┌─────────────────────────────────────────┐                   │
│  │  VISION LLM (gpt-4.1-mini)              │                   │
│  │  • Параллельная обработка страниц       │                   │
│  │  • Retry: 3 попытки + exponential backoff│                   │
│  │  • Fallback: текст об ошибке            │                   │
│  └─────────────────────────────────────────┘                   │
│         │                                                       │
│         ▼                                                       │
│  ┌─────────────────────────────────────────┐                   │
│  │  АГРЕГАЦИЯ                               │                   │
│  │  • Merge JSON со всех страниц           │                   │
│  │  • Обработка чанков на границах         │                   │
│  │  • Валидация полноты                    │                   │
│  └─────────────────────────────────────────┘                   │
│         │                                                       │
│         ▼                                                       │
│  ВЫХОД: Dict[str, str] — чанки для LLM-аудита                  │
│                                                                 │
└─────────────────────────────────────────────────────────────────┘
```

---

## 2. Зависимости и инструменты

### 2.1 Конвертация DOCX → PDF

#### LibreOffice Headless (рекомендуется)

**Команда:**
```bash
# macOS / Linux
libreoffice --headless --convert-to pdf --outdir /tmp document.docx

# или soffice
soffice --headless --convert-to pdf --outdir /tmp document.docx
```

**Преимущества:**
- Высокая точность конвертации (использует тот же движок, что и LibreOffice Writer)
- Бесплатный, open-source
- Работает на всех платформах

**Ограничения:**
- Не thread-safe (документы конвертируются по одному)
- Требует установки LibreOffice

**Альтернативы:**

| Инструмент | Описание | Рекомендация |
|------------|----------|--------------|
| `docx2pdf` | Python-обёртка, использует MS Word (Windows/macOS) или LibreOffice (Linux) | Хороший вариант для простых случаев |
| `unoconv` | **Устарел**, заменён на `unoserver` | Не использовать |
| `unoserver` | Современная замена unoconv | Для продакшн-серверов |
| `pandoc` | Универсальный конвертер | Только для простых документов |

#### Установка LibreOffice

```bash
# macOS
brew install --cask libreoffice

# Ubuntu/Debian
sudo apt-get install libreoffice-core libreoffice-writer

# Docker (Alpine)
apk add libreoffice
```

### 2.2 Конвертация PDF → PNG

#### pdf2image + poppler

```python
from pdf2image import convert_from_path

# Конвертация PDF в список PIL Image
images = convert_from_path(
    pdf_path,
    dpi=200,                    # Оптимальный DPI для Vision LLM
    fmt='png',                  # Формат выходных изображений
    thread_count=4,             # Параллельность
    use_pdftocairo=True,        # Может быть быстрее
    output_folder='/tmp/pages'  # Для больших PDF
)
```

**Оптимальный DPI:**

| DPI | Качество | Размер файла | Рекомендация |
|-----|----------|--------------|--------------|
| 90 | Низкое | Маленький | Плохо для OCR |
| 150 | Среднее | Средний | Минимум для Vision |
| 200 | Хорошее | Средний | **Рекомендуется** |
| 300 | Высокое | Большой | Для мелкого текста |

> **Важно:** GPT-4.1-mini при обработке изображений ресайзит их до 768px по короткой стороне. При DPI=90 это даёт ~90 DPI после ресайза, что приводит к ошибкам OCR. При DPI=200 после ресайза остаётся ~210 DPI — достаточно для качественного распознавания.

#### Установка poppler

```bash
# macOS
brew install poppler

# Ubuntu/Debian
sudo apt-get install poppler-utils

# Docker (Alpine)
apk add poppler-utils
```

---

## 3. Vision API (gpt-4.1-mini)

### 3.1 Формат запроса

```python
import base64
from openai import OpenAI

def encode_image_base64(image_path: str) -> str:
    """Кодирует изображение в base64."""
    with open(image_path, 'rb') as f:
        return base64.standard_b64encode(f.read()).decode('utf-8')

def call_vision_api(
    client: OpenAI,
    image_path: str,
    system_prompt: str,
    user_prompt: str,
    model: str = "gpt-4.1-mini",
    detail: str = "high",  # low | high | auto
    max_tokens: int = 4096
) -> str:
    """
    Вызывает Vision API для извлечения текста из изображения.

    Args:
        client: OpenAI клиент
        image_path: путь к изображению
        system_prompt: системный промпт
        user_prompt: пользовательский промпт
        model: модель (gpt-4.1-mini, gpt-4.1, gpt-4o)
        detail: уровень детализации (low=85 токенов, high=варьируется)
        max_tokens: максимум токенов в ответе

    Returns:
        Текст ответа от модели
    """
    # Кодируем изображение
    base64_image = encode_image_base64(image_path)

    # Определяем MIME тип
    ext = image_path.lower().split('.')[-1]
    mime_type = {
        'png': 'image/png',
        'jpg': 'image/jpeg',
        'jpeg': 'image/jpeg',
        'gif': 'image/gif',
        'webp': 'image/webp'
    }.get(ext, 'image/png')

    # Формируем сообщение
    messages = [
        {"role": "system", "content": system_prompt},
        {
            "role": "user",
            "content": [
                {"type": "text", "text": user_prompt},
                {
                    "type": "image_url",
                    "image_url": {
                        "url": f"data:{mime_type};base64,{base64_image}",
                        "detail": detail
                    }
                }
            ]
        }
    ]

    response = client.chat.completions.create(
        model=model,
        messages=messages,
        max_tokens=max_tokens,
        temperature=0.0
    )

    return response.choices[0].message.content
```

### 3.2 Параметр detail

| Значение | Описание | Токены |
|----------|----------|--------|
| `low` | Изображение 512x512, фиксированная стоимость | 85 токенов |
| `high` | Высокое разрешение, масштабирование в 2048x2048 | 765-1105+ токенов |
| `auto` | Модель выбирает сама | Варьируется |

**Рекомендация:** Использовать `high` для документов, `low` для предварительного просмотра.

### 3.3 Стоимость

| Модель | Input (за 1M токенов) | Output (за 1M токенов) |
|--------|----------------------|------------------------|
| gpt-4.1-mini | $0.15 | $0.60 |
| gpt-4.1 | $2.00 | $8.00 |
| gpt-4o | $2.50 | $10.00 |

**Для gpt-4.1-mini:** image токены умножаются на 1.62 для получения текстовых токенов.

**Пример расчёта стоимости документа (5 страниц, high detail):**
```
5 страниц × 800 image токенов × 1.62 = 6,480 input токенов
6,480 / 1,000,000 × $0.15 = $0.001 за документ

+ output (примерно 2000 токенов на документ)
2,000 / 1,000,000 × $0.60 = $0.0012

Итого: ~$0.002 за документ (5 страниц)
```

### 3.4 Лимиты

- **Максимальный размер изображения:** 20MB
- **Контекстное окно gpt-4.1-mini:** 1M токенов (можно отправить много изображений)
- **Поддерживаемые форматы:** PNG, JPEG, GIF, WEBP

---

## 4. Обработка ошибок и Fallback

### 4.1 Retry логика с Exponential Backoff

```python
from tenacity import (
    retry,
    stop_after_attempt,
    wait_exponential,
    retry_if_exception_type
)
from openai import RateLimitError, APIConnectionError, APITimeoutError

@retry(
    retry=retry_if_exception_type((RateLimitError, APIConnectionError, APITimeoutError)),
    stop=stop_after_attempt(3),
    wait=wait_exponential(multiplier=1, min=2, max=30)
)
def call_vision_with_retry(client, image_path, prompts):
    """
    Вызов Vision API с автоматическим retry.

    Exponential backoff: 2s → 4s → 8s → ... до 30s max
    """
    return call_vision_api(client, image_path, **prompts)
```

### 4.2 Fallback стратегия

```python
from dataclasses import dataclass
from typing import Optional
from enum import Enum

class ExtractionMethod(Enum):
    VISION_LLM = "vision_llm"
    ERROR = "error"

@dataclass
class PageExtractionResult:
    page_num: int
    text: str
    method: ExtractionMethod
    error: Optional[str] = None

def extract_page_with_fallback(
    image_path: str,
    page_num: int,
    client: OpenAI,
    prompts: dict
) -> PageExtractionResult:
    """
    Извлечение текста со страницы.

    Порядок:
    1. Vision LLM (gpt-4.1-mini) — 3 попытки с exponential backoff
    2. Текст об ошибке (если все попытки исчерпаны)
    """
    try:
        text = call_vision_with_retry(client, image_path, prompts)
        return PageExtractionResult(
            page_num=page_num,
            text=text,
            method=ExtractionMethod.VISION_LLM
        )
    except Exception as e:
        # Все retry исчерпаны — возвращаем ошибку
        return PageExtractionResult(
            page_num=page_num,
            text=f"[ОШИБКА ИЗВЛЕЧЕНИЯ СТРАНИЦЫ {page_num}]",
            method=ExtractionMethod.ERROR,
            error=str(e)
        )
```

---

## 5. Параллельная обработка

### 5.1 Asyncio + Semaphore

```python
import asyncio
from typing import List

async def process_pages_parallel(
    image_paths: List[str],
    client: OpenAI,
    prompts: dict,
    max_concurrent: int = 5
) -> List[PageExtractionResult]:
    """
    Параллельная обработка страниц с ограничением конкурентности.

    Args:
        image_paths: список путей к изображениям страниц
        client: OpenAI клиент
        prompts: словарь промптов
        max_concurrent: максимум одновременных запросов

    Returns:
        Список результатов извлечения
    """
    semaphore = asyncio.Semaphore(max_concurrent)

    async def process_one(page_num: int, image_path: str):
        async with semaphore:
            # Оборачиваем синхронный вызов в executor
            loop = asyncio.get_event_loop()
            result = await loop.run_in_executor(
                None,
                extract_page_with_fallback,
                image_path,
                page_num,
                client,
                prompts
            )
            return result

    # Запускаем все задачи параллельно
    tasks = [
        process_one(i, path)
        for i, path in enumerate(image_paths, 1)
    ]

    results = await asyncio.gather(*tasks)
    return list(results)
```

### 5.2 ThreadPoolExecutor (альтернатива)

```python
from concurrent.futures import ThreadPoolExecutor, as_completed

def process_pages_threaded(
    image_paths: List[str],
    client: OpenAI,
    prompts: dict,
    max_workers: int = 5
) -> List[PageExtractionResult]:
    """
    Обработка страниц через ThreadPool.

    Проще чем asyncio, но менее эффективно для I/O-bound задач.
    """
    results = []

    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        futures = {
            executor.submit(
                extract_page_with_fallback,
                path, i, client, prompts
            ): i
            for i, path in enumerate(image_paths, 1)
        }

        for future in as_completed(futures):
            results.append(future.result())

    # Сортируем по номеру страницы
    return sorted(results, key=lambda x: x.page_num)
```

### 5.3 Рекомендации по конкурентности

| Параметр | Рекомендация | Обоснование |
|----------|--------------|-------------|
| max_concurrent | 5-10 | Баланс скорости и rate limits |
| Rate limit | 20 RPS для gpt-4.1-mini | Может быть выше для платных планов |
| Документов в параллель | 3-5 | Для пакетной обработки |

---

## 6. Агрегация результатов

### 6.1 Объединение чанков

```python
from typing import Dict, List, Any

def merge_page_results(
    results: List[PageExtractionResult],
    chunk_config: Dict[str, Any]
) -> Dict[str, str]:
    """
    Объединяет результаты со всех страниц в единый dict чанков.

    Args:
        results: результаты извлечения страниц
        chunk_config: конфигурация чанков для данного типа документа

    Returns:
        Dict с чанками вида {"шапка": "...", "текст_приказа": "...", ...}
    """
    # Объединяем весь текст
    full_text = "\n\n".join([r.text for r in sorted(results, key=lambda x: x.page_num)])

    # Теперь парсим по маркерам (используем существующую логику)
    chunks = extract_chunks_by_markers(full_text, chunk_config)

    return chunks

def extract_chunks_by_markers(
    text: str,
    config: Dict[str, Any]
) -> Dict[str, str]:
    """
    Извлекает чанки из текста по маркерам из конфигурации.

    Маркеры — это ключевые фразы, разделяющие документ на логические части.
    """
    markers = config.get("markers", {})
    chunks = {}

    # Простой алгоритм: ищем маркеры и режем текст
    positions = []
    for chunk_name, marker in markers.items():
        pos = text.lower().find(marker.lower())
        if pos != -1:
            positions.append((pos, chunk_name, marker))

    # Сортируем по позиции
    positions.sort(key=lambda x: x[0])

    # Формируем чанки
    for i, (pos, name, marker) in enumerate(positions):
        if i + 1 < len(positions):
            end_pos = positions[i + 1][0]
        else:
            end_pos = len(text)

        chunks[name] = text[pos:end_pos].strip()

    return chunks
```

### 6.2 Обработка чанков на границе страниц

**Проблема:** Чанк может начаться на одной странице и закончиться на другой. Vision LLM может разбить его неправильно.

**Решение:** Использовать "overlap" при обработке страниц:

```python
def process_with_overlap(pages: List[str], overlap_chars: int = 500):
    """
    Обрабатывает страницы с перекрытием для корректного захвата границ.

    На каждой странице (кроме первой) захватываем последние overlap_chars
    с предыдущей страницы для контекста.
    """
    results = []
    prev_tail = ""

    for i, page in enumerate(pages):
        context = prev_tail + page
        # Обрабатываем context, но берём только новую часть

        # Сохраняем хвост для следующей страницы
        prev_tail = page[-overlap_chars:] if len(page) > overlap_chars else page

    return results
```

### 6.3 Валидация полноты

```python
def validate_chunks(
    chunks: Dict[str, str],
    required_chunks: List[str]
) -> List[str]:
    """
    Проверяет что все обязательные чанки присутствуют и не пусты.

    Returns:
        Список отсутствующих/пустых чанков
    """
    missing = []

    for chunk_name in required_chunks:
        if chunk_name not in chunks:
            missing.append(f"{chunk_name}: отсутствует")
        elif not chunks[chunk_name].strip():
            missing.append(f"{chunk_name}: пустой")
        elif len(chunks[chunk_name]) < 50:  # Подозрительно короткий
            missing.append(f"{chunk_name}: подозрительно короткий ({len(chunks[chunk_name])} символов)")

    return missing
```

---

## 7. Оценка производительности и стоимости

### 7.1 Время обработки

| Этап | Время (5 страниц) | Примечание |
|------|-------------------|------------|
| DOCX → PDF | 2-5 сек | LibreOffice headless |
| PDF → PNG | 1-2 сек | pdf2image |
| Vision LLM | 3-8 сек | Параллельно, 5 страниц |
| Агрегация | <1 сек | В памяти |
| **Итого** | **8-16 сек** | Один документ |

### 7.2 Стоимость на документ

| Компонент | Стоимость | Примечание |
|-----------|-----------|------------|
| Vision input (5 стр) | ~$0.001 | 6,500 токенов |
| Vision output | ~$0.001 | 2,000 токенов |
| **Итого** | **~$0.002** | За 5-страничный документ |

**При обработке 1000 документов:** ~$2.00

### 7.3 Сравнение с текущим подходом

| Параметр | python-docx | Vision Pipeline |
|----------|-------------|-----------------|
| Скорость | Мгновенно | 8-16 сек |
| Стоимость | $0 | ~$0.002/док |
| Точность | Высокая | Очень высокая |
| Сканы/PDF | Нет | Да |
| Таблицы | Ограничено | Хорошо |
| Сложность кода | Высокая | Низкая |

---

## 8. Интеграция с существующим кодом

### 8.1 Совместимость формата чанков

Vision Pipeline должен возвращать **тот же формат**, что и существующие парсеры:

```python
# Существующий формат (parser_prikaz_ic_docs.py)
result = {
    "имя_файла": "document.docx",
    "путь": "/path/to/document.docx",
    "шапка": "...",
    "текст_приказа": "...",
    "приложение_1_к_приказу": "...",
    "приложение_2_к_приказу": "...",
    "приложение_1_к_регламенту": "...",
    "лист_ознакомления": "..."
}

# Vision Pipeline должен вернуть такой же формат
```

### 8.2 Интерфейс VisionParser

```python
class VisionParser:
    """
    Универсальный Vision-парсер документов.

    Заменяет все существующие parser_*.py одним модулем.
    """

    def __init__(
        self,
        chunk_config_path: str,
        model: str = "gpt-4.1-mini",
        max_concurrent: int = 5,
        dpi: int = 200
    ):
        """
        Args:
            chunk_config_path: путь к JSON с конфигурацией чанков
            model: модель Vision LLM
            max_concurrent: параллельность
            dpi: разрешение конвертации
        """
        self.config = self._load_config(chunk_config_path)
        self.model = model
        self.max_concurrent = max_concurrent
        self.dpi = dpi
        self.client = OpenAI()

    def parse(self, document_path: str) -> Dict[str, str]:
        """
        Парсит документ (DOCX, PDF или изображение).

        Returns:
            Dict с чанками
        """
        # Определяем тип файла
        ext = Path(document_path).suffix.lower()

        if ext == '.docx':
            images = self._convert_docx_to_images(document_path)
        elif ext == '.pdf':
            images = self._convert_pdf_to_images(document_path)
        elif ext in ['.png', '.jpg', '.jpeg']:
            images = [document_path]
        else:
            raise ValueError(f"Неподдерживаемый формат: {ext}")

        # Извлекаем текст со всех страниц
        results = asyncio.run(
            process_pages_parallel(images, self.client, self.config['prompts'])
        )

        # Агрегируем в чанки
        chunks = merge_page_results(results, self.config)

        # Добавляем метаданные
        chunks["имя_файла"] = Path(document_path).name
        chunks["путь"] = str(Path(document_path).absolute())

        return chunks
```

---

## 9. Выводы и рекомендации

### 9.1 Ключевые выводы

1. **Vision Pipeline — viable решение** для замены python-docx парсеров
2. **gpt-4.1-mini** — оптимальный баланс цены и качества
3. **DPI 200** — достаточно для качественного OCR
4. **Параллельность 5-10** — оптимально для API
5. **Retry 3 попытки** — достаточно для надёжности

### 9.2 Рекомендации

| Приоритет | Рекомендация |
|-----------|--------------|
| Высокий | Начать с одной итерации (седьмая_итерация) |
| Высокий | Создать единый модуль vision_parser |
| Средний | Добавить кэширование конвертаций |
| Средний | Метрики и мониторинг |
| Низкий | A/B тестирование vs python-docx |

### 9.3 Риски

| Риск | Вероятность | Митигация |
|------|-------------|-----------|
| Rate limits API | Средняя | Semaphore + backoff |
| Высокая стоимость | Низкая | ~$0.002/документ |
| Ошибки Vision API | Низкая | Retry 3 попытки + текст ошибки |
| Время обработки | Средняя | Кэширование, параллельность |

---

## Ссылки

- [OpenAI Vision API Documentation](https://platform.openai.com/docs/guides/images-vision)
- [pdf2image на PyPI](https://pypi.org/project/pdf2image/)
- [Tesseract OCR Documentation](https://tesseract-ocr.github.io/tessdoc/ImproveQuality.html)
- [Tenacity Retry Library](https://github.com/jd/tenacity)
- [LibreOffice Headless Mode](https://www.baeldung.com/linux/latex-doc-docx-pdf-conversion)
