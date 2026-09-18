#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Тесты добора фрагментов PDF-страницы: распознавание страницы целиком теряет отдельные блоки
(номер приказа справа от даты — обратная связь заказчика 25.08), layout их видит, и они
дописываются на своё место.
"""
from src.format_parsers.pdf._parsing import longest_token, merge_missing_fragments, text_signature


def test_signature_and_token():
    assert text_signature("№ ОД-26/66") == "од2666"
    assert text_signature("г. Москва") == "гмосква"
    assert longest_token("№ ОД-26/66") == ""          # слов от 4 символов нет
    assert longest_token("ИНН 7716737990") == "7716737990"


def test_missing_number_is_inserted_next_to_date():
    page = "ПРИКАЗ\n\n10.08.2026\nг. Москва\n\nО системе подачи ППУ"
    fragments = ["ПРИКАЗ", "10.08.2026", "No ОД-26/66", "г. Москва", "О системе подачи ППУ"]
    merged, recovered = merge_missing_fragments(page, fragments)
    assert recovered == ["No ОД-26/66"]
    lines = merged.split("\n")
    assert lines.index("No ОД-26/66") == lines.index("10.08.2026") + 1
    assert lines.index("No ОД-26/66") < lines.index("г. Москва")


def test_present_fragments_are_not_duplicated():
    page = "ПРИКАЗ\n10.08.2026\nг. Москва"
    merged, recovered = merge_missing_fragments(page, ["ПРИКАЗ", "10.08.2026", "г. Москва"])
    assert recovered == [] and merged == page


def test_differently_recognized_fragment_is_not_duplicated():
    """Кроп прочитан иначе («UNH» вместо «ИНН»), но число совпадает — фрагмент не дублируется."""
    page = "ИНН 7716737990\nКПП 772901001"
    merged, recovered = merge_missing_fragments(page, ["UNH 7716737990", "KПП 772901001"])
    assert recovered == [] and merged == page


def test_short_fragments_are_skipped():
    """Совсем короткие подписи («М.П.») не добираются — риск дублей выше пользы."""
    page = "Генеральный директор\nСаяпин Е.В."
    merged, recovered = merge_missing_fragments(page, ["М.П.", "Саяпин Е.В."])
    assert recovered == [] and merged == page


def test_fragment_before_any_anchor_goes_first():
    page = "Текст страницы"
    merged, recovered = merge_missing_fragments(page, ["Приложение №1 к приказу", "Текст страницы"])
    assert recovered == ["Приложение №1 к приказу"]
    assert merged.split("\n")[0] == "Приложение №1 к приказу"
