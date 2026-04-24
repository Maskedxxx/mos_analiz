# START_MODULE_CONTRACT
# PURPOSE: Парсер doc_type «График обхода ОМ» (xlsx). Читает известные листы/ячейки и вытаскивает два семантических поля: шапку (название компании + ссылка на приказ) и перечень «ФИО — должность».
# INPUTS: Путь к .xlsx файлу.
# OUTPUTS: `GrafikObhodDocument` с полями `шапка` и `фио_должности`.
# KEYWORDS: grafik-obhod, xlsx, structured-parsing, secondary-file.
# LINKS: main.py::SECONDARY_PARSER_DISPATCH, doc_configs/*/config.json (secondary_file.parser="grafik_obhod").
# RATIONALE: Используется как secondary-парсер в связке с первичным документом (например, «Приказ об ответственных»). Выход уже структурирован под валидаторы — не нужна дальнейшая нарезка.
# END_MODULE_CONTRACT

from __future__ import annotations

# START_IMPORTS
import re
from typing import TypedDict

from openpyxl import load_workbook
# END_IMPORTS


# START_CONTRACT
# PURPOSE: Контракт возвращаемого значения парсера. Поля совпадают со scope-именами, ожидаемыми валидаторами.
# INPUTS: —
# OUTPUTS: TypedDict для IDE-автодополнения и статической проверки.
# KEYWORDS: typeddict, contract.
class GrafikObhodDocument(TypedDict):
    """
    Назначение:
        Контракт возвращаемого значения `parse_grafik_obhod`.

    Поля:
        шапка: Склейка названия организации и ссылки на приложение к приказу.
        фио_должности: Список «ФИО — должность», по одной записи на строку.

    Логика:
        Русские имена полей сохранены намеренно — валидаторы ссылаются на эти scope-имена
        в своих `rules.json` как `<chunk_prefix>шапка`, `<chunk_prefix>фио_должности`.
        Переименование потребует согласованной правки и в doc_configs.
    """

    шапка: str
    фио_должности: str
# END_CONTRACT


# START_PARSE_GRAFIK_OBHOD
# PURPOSE: Основная функция парсинга xlsx «График обхода ОМ».
# INPUTS: Путь к xlsx.
# OUTPUTS: `GrafikObhodDocument`.
# KEYWORDS: parse, xlsx, openpyxl, regex.
def parse_grafik_obhod(xlsx_path: str) -> GrafikObhodDocument:
    """
    Назначение:
        Читает xlsx «График обхода ОМ» и извлекает два семантических поля —
        шапку и список ФИО/должностей сотрудников.

    Вход:
        xlsx_path: Путь к .xlsx-файлу.

    Выход:
        `GrafikObhodDocument` с непустыми (или пустыми — если не нашли) полями.

    Логика:
        1. Открывает книгу в data_only-режиме (вычисленные формулы).
        2. Берёт лист с именем `график`/`шаблон` (регистр не важен); если таких нет —
           первый лист книги.
        3. Сканирует первые 50 строк и 35 колонок, собирая по одной строке непустые
           значения ячеек (порядок сохраняется).
        4. Для каждой строки:
           - ищет организационно-правовую форму (ООО/ЗАО/АО/ПАО/ИП) → название компании;
           - ищет упоминание «приложение ... приказ» → ссылка на приказ;
           - если первая ячейка похожа на ФИО (по регуляркам) и вторая не пустая и
             не «должность» — добавляет в список сотрудников в формате `ФИО — должность`.
        5. Дедуплицирует сотрудников через `dict.fromkeys` (сохраняется порядок).
        6. Собирает результат в `GrafikObhodDocument`.
    """
    wb = load_workbook(xlsx_path, data_only=True)

    # Приоритетные имена листов — `график` или `шаблон`; иначе берём первый лист.
    sheet_name = None
    for name in wb.sheetnames:
        if name.lower() in ("график", "шаблон"):
            sheet_name = name
            break
    if not sheet_name:
        sheet_name = wb.sheetnames[0]
    ws = wb[sheet_name]

    company_name = ""
    prikaz_ref = ""
    employees: list[str] = []

    # Лимитируем диапазон сканирования — граф обхода всегда умещается в первые 50×35.
    for row_num in range(1, min(51, ws.max_row + 1)):
        row_values: list[str] = []
        for col_num in range(1, min(35, ws.max_column + 1)):
            cell = ws.cell(row=row_num, column=col_num)
            if cell.value:
                row_values.append(str(cell.value).strip())

        # Орг-правовая форма → название компании (берём первое найденное).
        for val in row_values:
            if re.search(r"\b(ООО|ЗАО|АО|ПАО|ИП)\b", val) and not company_name:
                company_name = val

        # Ссылка на приложение к приказу — переносы строк склеиваем в пробел.
        for val in row_values:
            if "приложение" in val.lower() and "приказ" in val.lower():
                prikaz_ref = val.replace("\n", " ").strip()

        # Строка с сотрудником: первая ячейка похожа на ФИО, вторая — должность.
        if len(row_values) >= 2:
            first_val = row_values[0]
            second_val = row_values[1]
            if (
                first_val.lower() not in ("фио", "график", "")
                and "график обхода" not in first_val.lower()
                and "приложение" not in first_val.lower()
            ):
                has_fio_with_initials = re.search(r"[А-ЯЁ][а-яё]+\s+[А-ЯЁ]\.?[А-ЯЁа-яё]?\.?", first_val)
                has_fio_two_words = re.search(r"[А-ЯЁ][а-яё]+\s+[А-ЯЁ][а-яё]+", first_val)
                if has_fio_with_initials or has_fio_two_words:
                    if second_val and second_val.lower() != "должность":
                        employees.append(f"{first_val} — {second_val}")

    return {
        "шапка": f"{company_name}\n{prikaz_ref}".strip(),
        "фио_должности": "\n".join(dict.fromkeys(employees)),
    }
# END_PARSE_GRAFIK_OBHOD
