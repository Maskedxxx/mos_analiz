#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Тесты парсера карточки проекта 2.4: поля читаются по подписям на листе (раскладка формы у
заказчика сдвинута относительно прежних адресов), а нераспознанная форма не выдаётся за
«все поля пустые» — поднимается ошибка (жалоба 15.09 по ООО ГК РИОН).
"""
import pytest
from openpyxl import Workbook

from conftest import PROJECT_ROOT
from src.doc_type_parsers.kartochka_proekta import parse_kartochka_main, parse_metodika

SAMPLE = PROJECT_ROOT / "tests" / "data" / "kartochka_proekta_2_4" / "sample.xlsx"


@pytest.fixture
def sample_or_skip():
    if not SAMPLE.exists():
        pytest.skip(f"нет фикстуры {SAMPLE.relative_to(PROJECT_ROOT)} — см. tests/data/README.md")
    return SAMPLE


def test_main_sheet_fields_are_read(sample_or_skip, tmp_path):
    """Шапка, обе секции, показатели и события читаются, а не остаются пустыми."""
    doc = parse_kartochka_main(sample_or_skip, tmp_path)
    header = doc["header"]
    assert header["org_name"], "организация не прочитана"
    assert header["signee_position"] and header["signee_name"] and header["signee_date"]
    assert header["has_utverzhday"] is True
    for field in ("clients", "perimeter", "owner", "boundaries", "leader", "team"):
        assert doc["section1"][field], f"секция 1: поле {field} не прочитано"
    assert doc["section2"]["key_risk"] and doc["section2"]["justification"]
    assert len(doc["indicators"]) >= 2
    first = doc["indicators"][0]
    assert first["name"] and first["unit"] and first["base_value"] is not None
    assert len(doc["events"]) >= 5
    assert any(event["start_date"] for event in doc["events"])


def test_metodika_indicators_are_read(sample_or_skip, tmp_path):
    """Лист методики: показатели с единицами, способом расчёта и источником данных."""
    doc = parse_metodika(sample_or_skip, tmp_path)
    assert len(doc["indicators"]) >= 2
    first = doc["indicators"][0]
    assert first["name"] and first["calc_method"] and first["data_source"]


def test_indicator_dates_are_dates_or_none(sample_or_skip, tmp_path):
    """Даты замеров — только даты: число из строки показателя датой не считается."""
    dates = parse_kartochka_main(sample_or_skip, tmp_path)["indicator_dates"]
    for value in dates.values():
        assert value is None or value[:4].isdigit(), f"не дата: {value!r}"


def test_unknown_form_raises(tmp_path):
    """Лист с нужным именем, но без подписей полей → ошибка, а не «все поля пустые»."""
    wb = Workbook()
    wb.active.title = "Карточка проекта"
    wb.active["A1"] = "произвольный текст"
    path = tmp_path / "chuzhaya.xlsx"
    wb.save(path)
    with pytest.raises(ValueError, match="не распознана"):
        parse_kartochka_main(path, tmp_path)


# --- Единицы измерения: сокращение в справочнике и слово в карточке — одна единица ---

def test_unit_synonyms_normalize():
    from src.doc_type_validators.kartochka_proekta import _normalize_unit
    assert _normalize_unit("минута") == _normalize_unit("мин")
    assert _normalize_unit("Часы") == _normalize_unit("час")
    assert _normalize_unit("Погонный метр") == _normalize_unit("погонный метр")
    assert _normalize_unit("шт.") == "шт"
    assert _normalize_unit("Метров погонных на человека в час") == "метров погонных на человека в час"


def test_rule_9_accepts_synonym_and_reports_unknown():
    """«минута» при справочном «мин» — не замечание; единицы вне справочника — замечание."""
    from src.doc_type_validators.kartochka_proekta import _rule_9_validate
    dropdown = {"categories": {"Время протекания процесса": ["мин", "часы"], "Выработка": ["шт/смена"]}}
    ok = _rule_9_validate({"indicators": [{"number": 1, "name": "ВПП", "unit": "минута"}]}, dropdown)
    assert ok["status"] == "PASS", ok["discrepancy"]
    bad = _rule_9_validate(
        {"indicators": [{"number": 1, "name": "Выработка", "unit": "Метров погонных на человека в час"}]}, dropdown
    )
    assert bad["status"] == "FAIL" and "не найдена в справочнике" in bad["discrepancy"]


def test_rule_10_compares_units_by_meaning():
    """Карточка «минута» и методика «мин» — одно и то же, расхождения нет."""
    from src.doc_type_validators.kartochka_proekta import _rule_10_validate
    kartochka = {"indicators": [{"name": "ВПП", "unit": "минута"}]}
    metodika = {"indicators": [{"name": "ВПП:", "unit": "мин"}]}
    assert _rule_10_validate(kartochka, metodika)["status"] == "PASS"
