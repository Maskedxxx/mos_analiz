#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Конвертер документов DOCX → PDF.

Использует LibreOffice headless для конвертации документов.
Поддерживает DOCX, DOC, ODT и прямой проброс PDF.
"""

import os
import subprocess
import tempfile
import shutil
from pathlib import Path
from typing import Optional


class ConversionError(Exception):
    """Ошибка конвертации документа."""
    pass


class DocumentConverter:
    """
    Конвертер документов в PDF с использованием LibreOffice.

    Поддерживаемые форматы:
    - .docx, .doc, .odt → конвертация через LibreOffice
    - .pdf → возврат исходного пути без конвертации

    Атрибуты:
        temp_dir: Директория для временных файлов
        libreoffice_path: Путь к исполняемому файлу LibreOffice
    """

    # Форматы, требующие конвертации
    CONVERTIBLE_EXTENSIONS = {'.docx', '.doc', '.odt', '.rtf'}

    def __init__(self, temp_dir: Optional[str] = None):
        """
        Инициализация конвертера.

        Args:
            temp_dir: Директория для временных файлов.
                     Если не указана, используется системная temp директория.
        """
        self.temp_dir = Path(temp_dir) if temp_dir else Path(tempfile.gettempdir())
        self.temp_dir.mkdir(parents=True, exist_ok=True)
        self._created_files: list = []  # Отслеживаем созданные файлы для cleanup

        # Определяем путь к LibreOffice
        self.libreoffice_path = self._find_libreoffice()

    def _find_libreoffice(self) -> str:
        """
        Находит путь к LibreOffice.

        Returns:
            Путь к исполняемому файлу LibreOffice

        Raises:
            ConversionError: Если LibreOffice не найден
        """
        # Типичные пути для разных ОС
        possible_paths = [
            # macOS
            "/Applications/LibreOffice.app/Contents/MacOS/soffice",
            # Linux
            "/usr/bin/soffice",
            "/usr/bin/libreoffice",
            "/usr/local/bin/soffice",
            # Windows (через PATH)
            "soffice",
        ]

        # Проверяем через which/where
        for cmd in ["soffice", "libreoffice"]:
            result = shutil.which(cmd)
            if result:
                return result

        # Проверяем известные пути
        for path in possible_paths:
            if os.path.exists(path):
                return path

        raise ConversionError(
            "LibreOffice не найден. Установите LibreOffice:\n"
            "  macOS: brew install --cask libreoffice\n"
            "  Linux: apt install libreoffice\n"
            "  Windows: https://www.libreoffice.org/download/"
        )

    def convert(self, input_path: str) -> str:
        """
        Конвертирует документ в PDF.

        Args:
            input_path: Путь к исходному документу

        Returns:
            Путь к PDF файлу:
            - Исходный путь, если файл уже PDF
            - Путь к временному PDF, если была конвертация

        Raises:
            FileNotFoundError: Если исходный файл не найден
            ConversionError: Если конвертация не удалась
        """
        input_file = Path(input_path)

        if not input_file.exists():
            raise FileNotFoundError(f"Файл не найден: {input_path}")

        ext = input_file.suffix.lower()

        # PDF — возвращаем как есть
        if ext == '.pdf':
            return str(input_file.absolute())

        # Проверяем, поддерживается ли формат
        if ext not in self.CONVERTIBLE_EXTENSIONS:
            raise ConversionError(
                f"Неподдерживаемый формат: {ext}. "
                f"Поддерживаются: {', '.join(self.CONVERTIBLE_EXTENSIONS)}, .pdf"
            )

        # Конвертируем через LibreOffice
        return self._convert_with_libreoffice(input_file)

    def _convert_with_libreoffice(self, input_file: Path) -> str:
        """
        Конвертирует документ в PDF через LibreOffice headless.

        Args:
            input_file: Путь к исходному файлу

        Returns:
            Путь к созданному PDF

        Raises:
            ConversionError: Если конвертация не удалась
        """
        # Создаём временную директорию для вывода
        output_dir = self.temp_dir / "vision_converter"

        # ВАЖНО: Удаляем старый PDF с таким же именем, чтобы избежать кэширования
        old_pdf = output_dir / f"{input_file.stem}.pdf"
        if old_pdf.exists():
            old_pdf.unlink()

        output_dir.mkdir(exist_ok=True)

        # Формируем команду LibreOffice
        # --headless: без GUI
        # --convert-to pdf: конвертация в PDF
        # --outdir: директория для результата
        cmd = [
            self.libreoffice_path,
            "--headless",
            "--convert-to", "pdf",
            "--outdir", str(output_dir),
            str(input_file.absolute())
        ]

        try:
            result = subprocess.run(
                cmd,
                capture_output=True,
                text=True,
                timeout=120  # 2 минуты максимум
            )

            if result.returncode != 0:
                raise ConversionError(
                    f"LibreOffice вернул ошибку: {result.stderr}"
                )

        except subprocess.TimeoutExpired:
            raise ConversionError(
                "Таймаут конвертации (>120 секунд). "
                "Возможно, документ слишком большой."
            )
        except FileNotFoundError:
            raise ConversionError(
                f"Не удалось запустить LibreOffice: {self.libreoffice_path}"
            )

        # Определяем путь к результату
        output_pdf = output_dir / f"{input_file.stem}.pdf"

        if not output_pdf.exists():
            raise ConversionError(
                f"PDF не создан. Вывод LibreOffice:\n"
                f"stdout: {result.stdout}\n"
                f"stderr: {result.stderr}"
            )

        # Запоминаем для cleanup
        self._created_files.append(str(output_pdf))

        return str(output_pdf)

    def cleanup(self, pdf_path: str = None) -> None:
        """
        Удаляет временные файлы.

        Args:
            pdf_path: Конкретный PDF для удаления.
                     Если не указан, удаляет все созданные файлы.
        """
        if pdf_path:
            # Удаляем конкретный файл
            path = Path(pdf_path)
            if path.exists() and str(path) in self._created_files:
                path.unlink()
                self._created_files.remove(str(path))
        else:
            # Удаляем все созданные файлы
            for file_path in self._created_files:
                path = Path(file_path)
                if path.exists():
                    path.unlink()
            self._created_files.clear()

    def __enter__(self):
        """Поддержка context manager."""
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        """Автоматическая очистка при выходе из context manager."""
        self.cleanup()
        return False


def convert_to_pdf(input_path: str, output_path: str = None) -> str:
    """
    Утилитарная функция для быстрой конвертации.

    Args:
        input_path: Путь к исходному документу
        output_path: Путь для результата (опционально)

    Returns:
        Путь к PDF файлу
    """
    converter = DocumentConverter()
    pdf_path = converter.convert(input_path)

    if output_path:
        shutil.copy(pdf_path, output_path)
        converter.cleanup(pdf_path)
        return output_path

    return pdf_path
