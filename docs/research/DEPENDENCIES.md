# Зависимости Universal Vision Pipeline

**Дата:** 2026-01-30
**Версия:** 1.0

---

## 1. Обзор зависимостей

### 1.1 Категории

| Категория | Пакеты | Назначение |
|-----------|--------|------------|
| Конвертация | LibreOffice, poppler, pdf2image | DOCX→PDF→PNG |
| Vision LLM | openai | API для gpt-4.1-mini |
| Image Processing | Pillow | Работа с изображениями |
| Retry Logic | tenacity | Exponential backoff |
| Async | aiohttp, asyncio | Параллельная обработка |

---

## 2. Python пакеты

### 2.1 requirements.txt (дополнение)

```txt
# === Vision Pipeline Dependencies ===

# Конвертация PDF в изображения
pdf2image>=1.17.0

# Retry логика с exponential backoff
tenacity>=8.0.0

# Обработка изображений
Pillow>=10.0.0

# Асинхронные HTTP запросы (опционально)
aiohttp>=3.9.0

# === Существующие зависимости (уже есть) ===
openai>=1.59.0
python-docx>=1.1.0
pandas>=2.0.0
openpyxl>=3.1.0
```

### 2.2 Установка Python пакетов

```bash
# Активировать виртуальное окружение
source /path/to/venv/bin/activate

# Установить все зависимости
pip install pdf2image tenacity Pillow aiohttp

# Или через requirements.txt
pip install -r requirements.txt

# Проверить установку
python -c "import pdf2image; import tenacity; import PIL; print('OK')"
```

---

## 3. Системные зависимости

### 3.1 LibreOffice

LibreOffice используется для конвертации DOCX → PDF в headless режиме.

#### macOS

```bash
# Установка через Homebrew
brew install --cask libreoffice

# Проверка установки
/Applications/LibreOffice.app/Contents/MacOS/soffice --version

# Симлинк для удобства (опционально)
sudo ln -s /Applications/LibreOffice.app/Contents/MacOS/soffice /usr/local/bin/libreoffice
```

#### Ubuntu/Debian

```bash
# Установка
sudo apt-get update
sudo apt-get install -y libreoffice-core libreoffice-writer

# Проверка
libreoffice --version
# или
soffice --version
```

#### Docker (Alpine)

```dockerfile
# В Dockerfile
RUN apk add --no-cache libreoffice

# Или для более компактного образа (только Writer)
RUN apk add --no-cache libreoffice-writer
```

#### Docker (Debian/Ubuntu)

```dockerfile
# В Dockerfile
RUN apt-get update && \
    apt-get install -y --no-install-recommends \
    libreoffice-core \
    libreoffice-writer \
    && rm -rf /var/lib/apt/lists/*
```

### 3.2 Poppler

Poppler — набор утилит для работы с PDF. Требуется для pdf2image.

#### macOS

```bash
# Установка через Homebrew
brew install poppler

# Проверка
pdftoppm -v
# pdftoppm version 24.xx.x
```

#### Ubuntu/Debian

```bash
# Установка
sudo apt-get update
sudo apt-get install -y poppler-utils

# Проверка
pdftoppm -v
```

#### Docker (Alpine)

```dockerfile
RUN apk add --no-cache poppler-utils
```

#### Docker (Debian/Ubuntu)

```dockerfile
RUN apt-get update && \
    apt-get install -y --no-install-recommends poppler-utils \
    && rm -rf /var/lib/apt/lists/*
```

---

## 4. Полный Dockerfile

### 4.1 Alpine-based (компактный)

```dockerfile
# Vision Pipeline - Alpine
FROM python:3.11-alpine

# Системные зависимости
RUN apk add --no-cache \
    # LibreOffice для DOCX→PDF
    libreoffice-writer \
    # Poppler для PDF→PNG
    poppler-utils \
    # Зависимости для Pillow
    jpeg-dev \
    zlib-dev

# Python зависимости
WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Код приложения
COPY . .

# Точка входа
CMD ["python", "run_audit.py"]
```

### 4.2 Debian-based (более совместимый)

```dockerfile
# Vision Pipeline - Debian
FROM python:3.11-slim

# Установка системных зависимостей
RUN apt-get update && apt-get install -y --no-install-recommends \
    # LibreOffice для DOCX→PDF
    libreoffice-core \
    libreoffice-writer \
    # Poppler для PDF→PNG
    poppler-utils \
    # Очистка
    && rm -rf /var/lib/apt/lists/*

# Python зависимости
WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Код приложения
COPY . .

# Healthcheck
HEALTHCHECK --interval=30s --timeout=10s \
    CMD python -c "import pdf2image; import openai; print('OK')"

# Точка входа
CMD ["python", "run_audit.py"]
```

---

## 5. Скрипт проверки окружения

### 5.1 check_environment.py

```python
#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Скрипт проверки окружения для Vision Pipeline.

Проверяет наличие всех необходимых зависимостей.
"""

import sys
import subprocess
import shutil
from pathlib import Path


def check_python_version():
    """Проверка версии Python."""
    version = sys.version_info
    if version.major < 3 or (version.major == 3 and version.minor < 9):
        return False, f"Python {version.major}.{version.minor} (требуется >= 3.9)"
    return True, f"Python {version.major}.{version.minor}.{version.micro}"


def check_python_package(package_name: str, import_name: str = None):
    """Проверка Python пакета."""
    import_name = import_name or package_name
    try:
        module = __import__(import_name)
        version = getattr(module, '__version__', 'unknown')
        return True, f"{package_name} {version}"
    except ImportError:
        return False, f"{package_name} не установлен"


def check_system_command(cmd: str, version_flag: str = "--version"):
    """Проверка системной команды."""
    if not shutil.which(cmd):
        return False, f"{cmd} не найден в PATH"

    try:
        result = subprocess.run(
            [cmd, version_flag],
            capture_output=True,
            text=True,
            timeout=10
        )
        output = result.stdout or result.stderr
        version = output.strip().split('\n')[0]
        return True, version
    except Exception as e:
        return False, f"Ошибка: {e}"


def check_libreoffice():
    """Проверка LibreOffice."""
    # Пробуем разные имена команд
    for cmd in ["libreoffice", "soffice"]:
        if shutil.which(cmd):
            return check_system_command(cmd)

    # macOS: проверяем в Applications
    macos_path = "/Applications/LibreOffice.app/Contents/MacOS/soffice"
    if Path(macos_path).exists():
        try:
            result = subprocess.run(
                [macos_path, "--version"],
                capture_output=True,
                text=True,
                timeout=10
            )
            return True, result.stdout.strip().split('\n')[0]
        except Exception as e:
            return False, f"Ошибка: {e}"

    return False, "LibreOffice не найден"


def main():
    """Главная функция проверки."""
    print("=" * 60)
    print("ПРОВЕРКА ОКРУЖЕНИЯ VISION PIPELINE")
    print("=" * 60)

    checks = [
        ("Python", lambda: check_python_version()),
        ("", None),  # Разделитель
        ("--- Python пакеты ---", None),
        ("openai", lambda: check_python_package("openai")),
        ("pdf2image", lambda: check_python_package("pdf2image")),
        ("tenacity", lambda: check_python_package("tenacity")),
        ("Pillow", lambda: check_python_package("Pillow", "PIL")),
        ("python-docx", lambda: check_python_package("python-docx", "docx")),
        ("", None),  # Разделитель
        ("--- Системные зависимости ---", None),
        ("LibreOffice", lambda: check_libreoffice()),
        ("Poppler (pdftoppm)", lambda: check_system_command("pdftoppm", "-v")),
    ]

    all_ok = True

    for name, check_func in checks:
        if check_func is None:
            print(name)
            continue

        ok, info = check_func()
        status = "OK" if ok else "FAIL"
        symbol = "[+]" if ok else "[-]"
        print(f"  {symbol} {name}: {info}")

        if not ok:
            all_ok = False

    print("=" * 60)

    if all_ok:
        print("Все проверки пройдены! Окружение готово.")
        return 0
    else:
        print("ВНИМАНИЕ: Некоторые зависимости отсутствуют!")
        print("Установите недостающие компоненты перед запуском.")
        return 1


if __name__ == "__main__":
    sys.exit(main())
```

### 5.2 Использование

```bash
# Запуск проверки
python check_environment.py

# Пример вывода:
# ============================================================
# ПРОВЕРКА ОКРУЖЕНИЯ VISION PIPELINE
# ============================================================
#   [+] Python: 3.11.5
#
# --- Python пакеты ---
#   [+] openai: 1.59.7
#   [+] pdf2image: 1.17.0
#   [+] tenacity: 8.2.3
#   [+] Pillow: 10.2.0
#   [+] python-docx: 1.1.2
#
# --- Системные зависимости ---
#   [+] LibreOffice: LibreOffice 7.6.4.1
#   [+] Poppler (pdftoppm): pdftoppm version 24.02.0
# ============================================================
# Все проверки пройдены! Окружение готово.
```

---

## 6. Установка на разных платформах

### 6.1 macOS (полная установка)

```bash
#!/bin/bash
# install_macos.sh - Установка всех зависимостей на macOS

# Homebrew (если не установлен)
if ! command -v brew &> /dev/null; then
    /bin/bash -c "$(curl -fsSL https://raw.githubusercontent.com/Homebrew/install/HEAD/install.sh)"
fi

# Системные зависимости
brew install --cask libreoffice
brew install poppler

# Python зависимости
pip install pdf2image tenacity Pillow aiohttp

# Проверка
python check_environment.py
```

### 6.2 Ubuntu/Debian (полная установка)

```bash
#!/bin/bash
# install_ubuntu.sh - Установка всех зависимостей на Ubuntu/Debian

# Обновление пакетов
sudo apt-get update

# Системные зависимости
sudo apt-get install -y \
    libreoffice-core \
    libreoffice-writer \
    poppler-utils

# Python зависимости
pip install pdf2image tenacity Pillow aiohttp

# Проверка
python check_environment.py
```

### 6.3 Docker Compose

```yaml
# docker-compose.yml
version: '3.8'

services:
  vision-pipeline:
    build:
      context: .
      dockerfile: Dockerfile
    environment:
      - OPENAI_API_KEY=${OPENAI_API_KEY}
    volumes:
      - ./documents:/app/documents:ro
      - ./results:/app/results
    command: python run_audit.py --use-vision
```

---

## 7. Версии и совместимость

### 7.1 Минимальные версии

| Компонент | Минимальная версия | Рекомендуемая |
|-----------|-------------------|---------------|
| Python | 3.9 | 3.11+ |
| LibreOffice | 7.0 | 7.5+ |
| Poppler | 22.0 | 24.0+ |
| openai | 1.0 | 1.59+ |
| pdf2image | 1.16 | 1.17+ |

### 7.2 Известные проблемы

| Проблема | Решение |
|----------|---------|
| LibreOffice не запускается в Docker | Установить шрифты: `apt-get install fonts-dejavu` |
| pdf2image: "Unable to get page count" | Обновить poppler |
| LibreOffice timeout | Увеличить timeout в subprocess |
