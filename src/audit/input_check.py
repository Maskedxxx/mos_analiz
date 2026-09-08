# START_MODULE_CONTRACT
# PURPOSE: Проверка загруженного файла ДО запуска аудита: пустой, повреждённый, защищённый паролем, не тот формат под маской расширения. Даёт человекочитаемую причину вместо PackageNotFoundError/BadZipFile из глубины парсера (аудит устойчивости, находки 2.4c, 2.4d, 2.4i).
# INPUTS: путь к файлу (расширение берётся из имени).
# OUTPUTS: None — файл пригоден; ValueError с текстом для пользователя — нет.
# KEYWORDS: input-check, upload, zip, ooxml, encrypted, pdf.
# LINKS: src/api/server.py (start_audit, start_cross_audit), main.py (CLI).
# RATIONALE: Проверяется ровно то, что ломает парсеры: размер 0, OLE-контейнер (пароль/старый .doc), не-zip, отсутствие корневой части OOXML, сигнатура PDF. Содержимое документа здесь не анализируется.
# END_MODULE_CONTRACT

from __future__ import annotations

# START_IMPORTS
import zipfile
from pathlib import Path
# END_IMPORTS


# START_CONSTANTS
# Корневая часть OOXML-пакета по расширению (+ имя формата в творительном и именительном падеже для текстов).
_OOXML_ROOT = {
    ".docx": ("word/document.xml", "документом Word (.docx)", "документ Word"),
    ".pptx": ("ppt/presentation.xml", "презентацией PowerPoint (.pptx)", "презентация PowerPoint"),
    ".xlsx": ("xl/workbook.xml", "таблицей Excel (.xlsx)", "таблица Excel"),
}
# Какой формат на самом деле лежит в zip — для подсказки «это таблица, а не документ».
_OOXML_KIND = {"word/document.xml": "документ Word", "ppt/presentation.xml": "презентация PowerPoint", "xl/workbook.xml": "таблица Excel"}
_OLE_MAGIC = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"  # OLE2: зашифрованный OOXML (EncryptedPackage) или старый .doc/.xls
_PDF_MAGIC = b"%PDF"
# END_CONSTANTS


# START_CHECK_INPUT_FILE
def check_input_file(path: Path) -> None:
    """
    Назначение:
        Убедиться, что файл можно отдать парсеру, и объяснить пользователю, если нельзя.

    Вход:
        path: путь к загруженному файлу; формат определяется по расширению.

    Выход:
        None, если проверка пройдена. ValueError с человекочитаемым текстом — иначе.

    Логика:
        1. Файл существует и не пустой.
        2. Для .docx/.pptx/.xlsx: OLE-заголовок → «защищён паролем или старый формат»;
           не zip → «повреждён»; в zip нет корневой части нужного формата → «внутри другой формат».
        3. Для .pdf: сигнатура %PDF.
        Прочие расширения не проверяются (их отсеивает список допустимых у типа).
    """
    path = Path(path)
    if not path.exists():
        raise ValueError(f"Файл не найден: {path}")
    ext = path.suffix.lower()
    if path.stat().st_size == 0:
        raise ValueError("Файл пустой (0 байт) — загрузите заполненный документ.")
    with open(path, "rb") as f:
        head = f.read(8)
    if ext in _OOXML_ROOT:
        root_part, human, expected = _OOXML_ROOT[ext]
        if head.startswith(_OLE_MAGIC):
            raise ValueError(
                f"Файл защищён паролем или сохранён в старом формате (.doc/.xls). "
                f"Снимите пароль или пересохраните как {ext} и загрузите снова."
            )
        if not zipfile.is_zipfile(path):
            raise ValueError(f"Файл повреждён или не является {human}. Пересохраните документ и загрузите снова.")
        try:
            with zipfile.ZipFile(path) as z:
                names = set(z.namelist())
        except zipfile.BadZipFile:
            raise ValueError(f"Файл повреждён или не является {human}. Пересохраните документ и загрузите снова.")
        if root_part not in names:
            actual = next((kind for part, kind in _OOXML_KIND.items() if part in names), None)
            if actual:
                raise ValueError(f"Внутри файла — {actual}, а не {expected}. Загрузите файл нужного формата.")
            raise ValueError(f"Файл не является {human}. Загрузите файл нужного формата.")
    elif ext == ".pdf":
        if not head.startswith(_PDF_MAGIC):
            raise ValueError("Файл повреждён или не является PDF. Пересохраните документ и загрузите снова.")
# END_CHECK_INPUT_FILE
