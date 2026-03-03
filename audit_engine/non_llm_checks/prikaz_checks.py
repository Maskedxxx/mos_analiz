#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Детерминированные проверки для приказов (prikaz_comp_ppu, prikaz_ppu и др.).

Правило 4: наличие слова ПРИКАЗ и номера приказа.
Правило 5: наличие города и даты приказа.
Правило 7: наличие должности и ФИО подписанта.

Переведены с LLM на non-LLM из-за систематических ошибок:
- LLM путает scope (проверяет номер вместо даты)
- LLM отвергает garbled OCR номера как «некорректные»
- LLM не распознаёт словесный формат даты
- LLM считает фамилии на -ович отчествами (Александрович М.Д.)
"""

import re
from typing import Any, Dict, List

from .registry import register


def _get_scopes_text(target_doc: Dict[str, Any], scopes: List[str]) -> str:
    """Собирает текст из нескольких scope в одну строку."""
    parts = []
    for scope in scopes:
        val = target_doc.get(scope, "")
        if val:
            parts.append(str(val))
    return "\n".join(parts)


def check_prikaz_and_number(
    target_doc: Dict[str, Any],
    config: Any,
    rule_index: int = 4,
    rule_title: str = "Проверка наличия слова ПРИКАЗ и номера приказа.",
    scopes: List[str] = None
) -> List[Dict[str, Any]]:
    """
    Проверяет наличие слова ПРИКАЗ и номера приказа.

    Допустимые формы: ПРИКАЗ, Приказ, П Р И К А З
    Номер: символ № + любой непустой текст (кроме подчёркиваний)
    scopes: список scope для извлечения текста (по умолчанию заголовок_город + номер_дата)
    """
    if scopes is None:
        scopes = ["заголовок_город", "номер_дата"]
    text = _get_scopes_text(target_doc, scopes)
    violations = []

    # Проверка слова ПРИКАЗ
    has_prikaz = bool(re.search(
        r'П\s*Р\s*И\s*К\s*А\s*З|Приказ',
        text,
        re.IGNORECASE
    ))

    if not has_prikaz:
        violations.append({
            "rule_index": rule_index,
            "rule_title": rule_title,
            "Целевой документ": "отсутствует",
            "Различие": "Отсутствует слово «ПРИКАЗ»"
        })

    # Проверка номера: ищем № с непустым текстом после него
    number_match = re.search(r'№\s*(.+)', text)
    if number_match:
        # Текст после №
        after_sign = number_match.group(1).strip()
        # Убираем подчёркивания и пробелы — если осталось что-то, номер заполнен
        cleaned = re.sub(r'[_\s]', '', after_sign)
        # Проверяем «б/н», «б\н», «бн» — аббревиатура «без номера»
        is_bn = bool(re.match(r'^б[/\\]?н$', cleaned, re.IGNORECASE))
        if not cleaned or is_bn:
            violations.append({
                "rule_index": rule_index,
                "rule_title": rule_title,
                "Целевой документ": f"№ {after_sign}",
                "Различие": "Номер приказа не заполнен" + (" (б/н = без номера)" if is_bn else " (пустой или подчёркивания)")
            })
    else:
        violations.append({
            "rule_index": rule_index,
            "rule_title": rule_title,
            "Целевой документ": "отсутствует",
            "Различие": "Отсутствует символ № с номером приказа"
        })

    return violations


def check_city_and_date(
    target_doc: Dict[str, Any],
    config: Any,
    rule_index: int = 5,
    rule_title: str = "Проверка наличия города и даты приказа.",
    scopes: List[str] = None
) -> List[Dict[str, Any]]:
    """
    Проверяет наличие города и полной даты приказа.

    Город: «г. Москва», полный адрес с городом, «Москва»
    Дата: числовой (ДД.ММ.ГГГГ) или словесный (01 сентября 2025 г.)
    scopes: список scope для извлечения текста (по умолчанию заголовок_город + номер_дата)
    """
    if scopes is None:
        scopes = ["заголовок_город", "номер_дата"]
    text = _get_scopes_text(target_doc, scopes)
    violations = []

    # Проверка города
    # Принимаем город в ЛЮБОМ контексте: "г. Москва", "Москва", ФИАС-адрес с городом
    has_city = bool(re.search(
        r'г\.\s*[А-ЯЁ][а-яё]+|'           # г. Москва, г.Москва
        r'город\s+[А-ЯЁ]|'                  # город Москва
        r'Москв[аеы]|'                       # Москва в любом контексте (включая ФИАС)
        r'Санкт-Петербург',                  # Санкт-Петербург
        text
    ))

    if not has_city:
        violations.append({
            "rule_index": rule_index,
            "rule_title": rule_title,
            "Целевой документ": "отсутствует",
            "Различие": "Отсутствует город подписания (например «г. Москва»)"
        })

    # Проверка даты
    # Месяцы для словесного формата
    months = (
        r'январ[яь]|феврал[яь]|март[а]?|апрел[яь]|ма[йя]|'
        r'июн[яь]|июл[яь]|август[а]?|сентябр[яь]|октябр[яь]|'
        r'ноябр[яь]|декабр[яь]'
    )

    has_date = bool(re.search(
        r'\d{2}[.\s]+\d{2}[.\s]+\d{4}|'                     # ДД.ММ.ГГГГ
        r'[«"\'"]?\d{1,2}[»"\'"]?\s*(?:' + months + r')',    # ДД месяц (словесный)
        text,
        re.IGNORECASE
    ))

    # Проверяем что дата не является плейсхолдером (__.__.202_)
    if has_date:
        placeholder = re.search(r'_+\._+\._+', text)
        if placeholder:
            has_date = False

    if not has_date:
        violations.append({
            "rule_index": rule_index,
            "rule_title": rule_title,
            "Целевой документ": "отсутствует или неполная",
            "Различие": "Отсутствует полная дата приказа"
        })

    return violations


def check_signatory(
    target_doc: Dict[str, Any],
    config: Any,
    rule_index: int = 7,
    rule_title: str = "Проверка должности и ФИО подписанта.",
    scopes: List[str] = None
) -> List[Dict[str, Any]]:
    """
    Проверяет наличие должности и ФИО подписанта.

    Должность: любое словосочетание (директор, начальник, управляющий и др.)
    ФИО: слово + инициалы (Иванов А.В.) или инициалы + слово (А.В. Иванов)
    Переведено на non-LLM из-за систематической ошибки LLM с фамилиями на -ович
    """
    if scopes is None:
        scopes = ["должность_фио_подписанта"]
    text = _get_scopes_text(target_doc, scopes)
    violations = []

    if not text.strip():
        violations.append({
            "rule_index": rule_index,
            "rule_title": rule_title,
            "Целевой документ": "отсутствует",
            "Различие": "Блок подписанта пуст"
        })
        return violations

    # Ошибка парсинга: в scope подписанта попал весь документ.
    # Признак: содержит слово «ПРИКАЗЫВАЮ» (которого не бывает в подписи).
    # Подписант располагается ПОСЛЕ последнего нумерованного пункта.
    if re.search(r'ПРИКАЗЫВАЮ', text, re.IGNORECASE):
        # Разбиваем на абзацы (по пустым строкам)
        paragraphs = [p.strip() for p in re.split(r'\n\s*\n', text) if p.strip()]
        if paragraphs:
            last_block = paragraphs[-1]
            # Если последний абзац — нумерованный пункт, подписанта НЕТ
            if re.match(r'\d+\.\s', last_block):
                text = ''
            else:
                # Последний абзац не пункт — это возможный подписант
                text = last_block

    # Проверка ФИО: слово + И.О. или И.О. + слово
    # Паттерны: "Иванов А.В.", "А.В. Иванов", "Иванов А. В.", "А. В. Иванов"
    fio_pattern = (
        r'[А-ЯЁа-яё]{2,}\s+[А-ЯЁ]\s*\.\s*[А-ЯЁ]\s*\.|'  # Фамилия И.О.
        r'[А-ЯЁ]\s*\.\s*[А-ЯЁ]\s*\.\s*[А-ЯЁа-яё]{2,}'     # И.О. Фамилия
    )
    has_fio = bool(re.search(fio_pattern, text))

    # Плейсхолдеры ФИО — невалидно
    # НЕ включаем _{4,} — подчёркивания могут быть линией подписи рядом с реальным ФИО
    # (формат: «Директор __________________ И.О. Фамилия»)
    fio_placeholders = re.search(
        r'И\.О\.\s*Фамилия|Фамилия\s*И\.О\.|'
        r'(?<![А-ЯЁа-яё])ФИО(?![А-ЯЁа-яё])',
        text
    )
    if fio_placeholders:
        has_fio = False

    if not has_fio:
        violations.append({
            "rule_index": rule_index,
            "rule_title": rule_title,
            "Целевой документ": text.strip()[:100],
            "Различие": "ФИО подписанта не заполнено или содержит плейсхолдер"
        })

    # Проверка должности: любое осмысленное слово перед ФИО
    # Исключаем плейсхолдеры должности
    dolzhnost_placeholder = re.search(
        r'указать\s+наименование\s+должности|'
        r'\(должность\)',
        text, re.IGNORECASE
    )
    # Проверяем наличие хотя бы одного слова кроме ФИО
    # Должность = текст без ФИО и служебных символов
    text_no_fio = re.sub(fio_pattern, '', text)
    text_no_fio = re.sub(r'[_\-—/\\«»"\'"().\d]', ' ', text_no_fio)
    words = [w for w in text_no_fio.split() if len(w) >= 3 and w not in ('М.П.', 'МП')]
    has_dolzhnost = len(words) >= 1 and not dolzhnost_placeholder

    if not has_dolzhnost:
        violations.append({
            "rule_index": rule_index,
            "rule_title": rule_title,
            "Целевой документ": text.strip()[:100],
            "Различие": "Должность подписанта не указана"
        })

    return violations


def check_fio_after_marker(
    target_doc: Dict[str, Any],
    config: Any,
    rule_index: int,
    rule_title: str,
    marker_phrase: str,
    scopes: List[str] = None
) -> List[Dict[str, Any]]:
    """
    Проверяет наличие должности и ФИО после фразы-маркера в тексте приказа.

    Используется для проверки заполненности назначенных лиц (организатор, секретарь).
    Ищет marker_phrase → после неё должна быть должность + ФИО.
    """
    if scopes is None:
        scopes = ["текст_приказа"]
    text = _get_scopes_text(target_doc, scopes)
    violations = []

    # Ищем фразу-маркер (без учёта регистра)
    match = re.search(re.escape(marker_phrase), text, re.IGNORECASE)
    if not match:
        # Фраза-маркер не найдена — пропускаем (структура документа отличается)
        return violations

    # Берём текст после маркера до конца строки/пункта
    after_marker = text[match.end():]
    # Ограничиваем до следующего пункта (цифра+точка в начале строки)
    next_punkt = re.search(r'\n\d+\.', after_marker)
    if next_punkt:
        after_marker = after_marker[:next_punkt.start()]

    # Проверяем ФИО: Фамилия И.О. или И.О. Фамилия, или полное ФИО (Балашова Романа Анатольевича)
    fio_pattern_short = (
        r'[А-ЯЁа-яё]{2,}\s+[А-ЯЁ]\s*\.\s*[А-ЯЁ]\s*\.|'    # Фамилия И.О.
        r'[А-ЯЁ]\s*\.\s*[А-ЯЁ]\s*\.\s*[А-ЯЁа-яё]{2,}'      # И.О. Фамилия
    )
    fio_pattern_full = (
        r'[А-ЯЁ][а-яё]{2,}\s+[А-ЯЁ][а-яё]{2,}\s+[А-ЯЁ][а-яё]{2,}'  # Фамилия Имя Отчество (любые фамилии, включая нестандартные)
    )
    fio_pattern = fio_pattern_short + '|' + fio_pattern_full
    fio_match = re.search(fio_pattern, after_marker)
    has_fio = bool(fio_match)

    # Плейсхолдеры
    has_placeholder = bool(re.search(
        r'И\.О\.\s*Фамилия|ФИО|_{4,}|\(указать\)',
        after_marker, re.IGNORECASE
    ))

    # Проверка должности: между маркером и ФИО должно быть слово-должность
    # Пример нормы: "...на производственной площадке координатора проектов Фролову А.О."
    # Пример НЕ нормы: "...на производственной площадке Жукову Анжелику Анатольевну" (нет должности)
    has_dolzhnost = True
    if has_fio and fio_match:
        # Текст между концом маркера (начало after_marker) и началом ФИО
        text_before_fio = after_marker[:fio_match.start()].strip()
        # Убираем предлоги, союзы, знаки препинания — ищем содержательные слова
        # Убираем «на [производственной] площадке» — это не должность
        text_before_fio = re.sub(r'на\s+(?:производственной\s+)?площадке', '', text_before_fio, flags=re.IGNORECASE)
        words_before = [w for w in re.findall(r'[А-ЯЁа-яё]{3,}', text_before_fio)
                        if w.lower() not in ('на', 'по', 'для', 'при', 'из', 'от', 'над')]
        has_dolzhnost = len(words_before) >= 1

    if has_placeholder or not has_fio:
        violations.append({
            "rule_index": rule_index,
            "rule_title": rule_title,
            "Целевой документ": after_marker.strip()[:150],
            "Различие": "ФИО после маркера не заполнено" if not has_fio
                         else "Содержит плейсхолдер вместо реального ФИО"
        })
    elif not has_dolzhnost:
        violations.append({
            "rule_index": rule_index,
            "rule_title": rule_title,
            "Целевой документ": after_marker.strip()[:150],
            "Различие": "Должность перед ФИО не указана"
        })

    return violations


# === Регистрация для prikaz_comp_ppu ===

@register("prikaz_comp_ppu", 4)
def check_prikaz_number_comp_ppu(target_doc, config):
    """Правило #4: ПРИКАЗ + номер для Приказа о конкурсах ППУ."""
    return check_prikaz_and_number(target_doc, config, 4,
                                   "Проверка наличия слова ПРИКАЗ и номера приказа.")


@register("prikaz_comp_ppu", 5)
def check_city_date_comp_ppu(target_doc, config):
    """Правило #5: город + дата для Приказа о конкурсах ППУ."""
    return check_city_and_date(target_doc, config, 5,
                                "Проверка наличия города и даты приказа.")


@register("prikaz_comp_ppu", 7)
def check_signatory_comp_ppu(target_doc, config):
    """Правило #7: должность + ФИО подписанта для Приказа о конкурсах ППУ."""
    return check_signatory(target_doc, config, 7,
                           "Проверка должности и ФИО подписанта.")


# === Регистрация для prikaz_ppu ===

@register("prikaz_ppu", 4)
def check_prikaz_number_ppu(target_doc, config):
    """Правило #4: ПРИКАЗ + номер для Приказа о ППУ."""
    return check_prikaz_and_number(target_doc, config, 4,
                                   "Проверка наличия слова ПРИКАЗ и номера приказа.")


@register("prikaz_ppu", 5)
def check_city_date_ppu(target_doc, config):
    """Правило #5: город + дата для Приказа о ППУ."""
    return check_city_and_date(target_doc, config, 5,
                                "Проверка наличия города и даты приказа.")


@register("prikaz_ppu", 7)
def check_signatory_ppu(target_doc, config):
    """Правило #7: должность + ФИО подписанта для Приказа о ППУ."""
    return check_signatory(target_doc, config, 7,
                           "Проверка должности и ФИО подписанта.")


# === Регистрация для prikaz_vyhod ===

@register("prikaz_vyhod", 4)
def check_prikaz_number_vyhod(target_doc, config):
    """Правило #4: ПРИКАЗ + номер для Приказа о проведении выхода."""
    return check_prikaz_and_number(target_doc, config, 4,
                                   "Проверка наличия слова ПРИКАЗ и номера приказа")


@register("prikaz_vyhod", 5)
def check_city_date_vyhod(target_doc, config):
    """Правило #5: город + дата для Приказа о проведении выхода."""
    return check_city_and_date(target_doc, config, 5,
                                "Проверка наличия города и даты приказа")


@register("prikaz_vyhod", 7)
def check_signatory_vyhod(target_doc, config):
    """Правило #7: должность + ФИО подписанта для Приказа о проведении выхода."""
    return check_signatory(target_doc, config, 7,
                           "Проверка наличия должности и ФИО подписанта")


@register("prikaz_vyhod", 8)
def check_organizer_vyhod(target_doc, config):
    """Правило #8: ФИО организатора (п.1) для Приказа о проведении выхода."""
    return check_fio_after_marker(
        target_doc, config, 8,
        "Проверка заполнения ФИО организатора (п.1)",
        "назначить организатором проведения обхода")


@register("prikaz_vyhod", 9)
def check_secretary_vyhod(target_doc, config):
    """Правило #9: ФИО секретаря (п.2) для Приказа о проведении выхода."""
    return check_fio_after_marker(
        target_doc, config, 9,
        "Проверка заполнения ФИО секретаря (п.2)",
        "назначить секретарем проведения обхода")


# === Регистрация для prikaz_ic_potoka ===
# У этого типа scope "шапка" вместо ["заголовок_город", "номер_дата"]

@register("prikaz_ic_potoka", 4)
def check_prikaz_number_ic_potoka(target_doc, config):
    """Правило #4: ПРИКАЗ + номер для Приказа о создании ИЦ потока."""
    return check_prikaz_and_number(target_doc, config, 4,
                                   "Проверка слова ПРИКАЗ и номера.",
                                   scopes=["шапка"])


@register("prikaz_ic_potoka", 5)
def check_city_date_ic_potoka(target_doc, config):
    """Правило #5: город + дата для Приказа о создании ИЦ потока."""
    return check_city_and_date(target_doc, config, 5,
                                "Проверка города и даты приказа.",
                                scopes=["шапка"])


@register("prikaz_ic_potoka", 7)
def check_signatory_ic_potoka(target_doc, config):
    """Правило #7: должность + ФИО подписанта для Приказа о создании ИЦ потока."""
    return check_signatory(target_doc, config, 7,
                           "Проверка подписанта.",
                           scopes=["подписант"])


@register("prikaz_ic_potoka", 8)
def check_responsible_fio_ic_potoka(target_doc, config):
    """
    Правило #8: ФИО ответственных лиц для Приказа о создании ИЦ потока.

    Проверяет 4 позиции:
    A) текст_приказа — пункт с «ознакомить»/«ознакомление» → должность+ФИО
    B) приложение_2_регламент п.2.1 — ответственный за ИЦ → ФИО
    C) приложение_2_регламент п.2.2 — исполняющий обязанности → ФИО
    D) приложение_2_регламент п.2.3 — администратор → ФИО

    Переведено с LLM на non-LLM из-за систематической ошибки:
    LLM не распознаёт ФИО после «исполняющему обязанности» (позиция C).
    Также: context_filter LLM жёстко привязан к нумерации пунктов,
    non-LLM ищет семантически — устойчив к сдвигу нумерации.
    """
    violations = []
    rule_index = 8
    rule_title = "Проверка ФИО ответственных лиц."

    # Паттерн ФИО: Фамилия И.О. или И.О. Фамилия
    fio_pattern = (
        r'[А-ЯЁа-яё]{2,}\s+[А-ЯЁ]\s*\.\s*[А-ЯЁ]\s*\.|'  # Иванова А.В.
        r'[А-ЯЁ]\s*\.\s*[А-ЯЁ]\s*\.\s*[А-ЯЁа-яё]{2,}'    # А.В. Иванова
    )

    # === Позиция A: текст_приказа — пункт с «ознакомить»/«ознакомление» ===
    text_prikaz = target_doc.get("текст_приказа", "")
    pos_a_paragraph = ""
    for line in text_prikaz.split('\n'):
        if re.search(r'ознакомл|ознакомит', line, re.IGNORECASE):
            pos_a_paragraph = line.strip()
            break

    if not pos_a_paragraph:
        # Пункт с «ознакомлением» не найден — нарушение
        violations.append({
            "rule_index": rule_index,
            "rule_title": rule_title,
            "Целевой документ": text_prikaz.strip()[:200],
            "Различие": "Позиция A: пункт с «ознакомлением» отсутствует в тексте приказа"
        })
    elif not re.search(fio_pattern, pos_a_paragraph):
        # Пункт есть, но ФИО отсутствует
        violations.append({
            "rule_index": rule_index,
            "rule_title": rule_title,
            "Целевой документ": pos_a_paragraph[:200],
            "Различие": "Позиция A: в пункте с «ознакомлением» не указано ФИО ответственного"
        })

    # === Позиции B, C, D: приложение_2_регламент ===
    reglament = target_doc.get("приложение_2_регламент", "")
    # Если регламент отсутствует или пуст — проверяем только позицию A
    if not reglament or not reglament.strip() or '[НЕТ СТРАНИЦ' in reglament or '[Фильтр: не найдено]' in reglament:
        return violations

    # Вспомогательная функция: извлекает текст параграфа между двумя маркерами
    def extract_paragraph(text, start_re, end_re):
        start = re.search(start_re, text)
        if not start:
            return ""
        rest = text[start.end():]
        end = re.search(end_re, rest)
        if end:
            return rest[:end.start()].strip()
        return rest.strip()

    # Позиция B: п.2.1 — ответственный за работу ИЦ → должно быть ФИО
    text_21 = extract_paragraph(reglament, r'2\.1\.?\s', r'\n\s*2\.2')
    if text_21 and not re.search(fio_pattern, text_21):
        violations.append({
            "rule_index": rule_index,
            "rule_title": rule_title,
            "Целевой документ": f"п.2.1: {text_21[:200]}",
            "Различие": "Позиция B: в п.2.1 регламента не указано ФИО ответственного за ИЦ"
        })

    # Позиция C: п.2.2 — исполняющий обязанности → ФИО ПОСЛЕ этих слов
    text_22 = extract_paragraph(reglament, r'2\.2\.?\s', r'\n\s*2\.3')
    if text_22:
        io_match = re.search(r'исполняющ\w*\s+обязанност\w*', text_22, re.IGNORECASE)
        if io_match:
            # Проверяем ФИО ПОСЛЕ «исполняющему обязанности»
            after_io = text_22[io_match.end():]
            if not re.search(fio_pattern, after_io):
                violations.append({
                    "rule_index": rule_index,
                    "rule_title": rule_title,
                    "Целевой документ": f"п.2.2: {text_22[:200]}",
                    "Различие": "Позиция C: после «исполняющему обязанности» не указано ФИО"
                })

    # Позиция D: п.2.3 — администратор → ФИО
    text_23 = extract_paragraph(reglament, r'2\.3\.?\s', r'\n\s*2\.4')
    if text_23:
        # Ищем строку с «Администратор» и проверяем ФИО в ней
        for line in text_23.split('\n'):
            if re.search(r'[Аа]дминистратор', line):
                if not re.search(fio_pattern, line):
                    violations.append({
                        "rule_index": rule_index,
                        "rule_title": rule_title,
                        "Целевой документ": f"п.2.3: {line.strip()[:200]}",
                        "Различие": "Позиция D: для администратора ИЦ не указано ФИО"
                    })
                break

    return violations


# ============================================================================
# check_signatory_with_stamp — подпись + М.П./МП/ПЕЧАТЬ
# ============================================================================

def check_signatory_with_stamp(
    target_doc: Dict[str, Any],
    config: Any,
    rule_index: int = 9,
    rule_title: str = "Проверка подписи директора и М.П.",
    scopes: List[str] = None
) -> List[Dict[str, Any]]:
    """
    Проверяет наличие должности, ФИО и печати (М.П./МП/ПЕЧАТЬ) в подписанте.

    Переведено с LLM на non-LLM: gpt-4.1-mini не видит «М.П.» в собственном
    контексте (hallucination) и не распознаёт «МП» без точек как аналог «М.П.».
    """
    if scopes is None:
        scopes = ["подписант"]
    text = _get_scopes_text(target_doc, scopes)

    violations = []
    if not text or len(text.strip()) < 3:
        violations.append({
            "rule_index": rule_index,
            "rule_title": rule_title,
            "Целевой документ": "(текст подписанта отсутствует)",
            "Различие": "Отсутствует блок подписанта"
        })
        return violations

    # --- Проверка должности ---
    position_words = [
        'директор', 'руководитель', 'начальник', 'управляющ',
        'заместител', 'президент', 'председател'
    ]
    text_lower = text.lower()
    has_position = any(w in text_lower for w in position_words)
    if not has_position:
        violations.append({
            "rule_index": rule_index,
            "rule_title": rule_title,
            "Целевой документ": text.strip()[:200],
            "Различие": "Отсутствует должность подписанта (Генеральный директор и т.п.)"
        })

    # --- Проверка ФИО ---
    fio_pattern = r'[А-ЯЁ][а-яё]+\s+[А-ЯЁ]\.[А-ЯЁ]\.|[А-ЯЁ]\.[А-ЯЁ]\.\s*[А-ЯЁ][а-яё]+'
    placeholder_pattern = r'_{4,}|И\.О\.\s*Фамилия|Фамилия\s*И\.О\.'
    has_fio = bool(re.search(fio_pattern, text))
    has_placeholder = bool(re.search(placeholder_pattern, text))
    if not has_fio or has_placeholder:
        violations.append({
            "rule_index": rule_index,
            "rule_title": rule_title,
            "Целевой документ": text.strip()[:200],
            "Различие": "ФИО подписанта не заполнено (плейсхолдер или отсутствует)"
        })

    # --- Проверка печати (М.П. / МП / ПЕЧАТЬ) ---
    # М.П., М. П., МП (с точками или без), ПЕЧАТЬ, печать
    stamp_pattern = r'М\.?\s*П\.?|ПЕЧАТЬ|печать'
    has_stamp = bool(re.search(stamp_pattern, text))
    if not has_stamp:
        violations.append({
            "rule_index": rule_index,
            "rule_title": rule_title,
            "Целевой документ": text.strip()[:200],
            "Различие": "Отсутствует указание на печать (М.П., МП или ПЕЧАТЬ)"
        })

    return violations


# ============================================================================
# Регистрация для prikaz_formirovanie_po
# ============================================================================

@register("prikaz_formirovanie_po", 3)
def check_rekvizity_formirovanie_po(target_doc, config):
    """
    Правило #3: реквизиты приказа — юрлицо, дата, номер, город.

    Проверяет наличие 4 обязательных реквизитов в шапке документа:
    1) Юрлицо: ООО/ЗАО/АО/ПАО + наименование
    2) Дата: числовой (ДД.ММ.ГГГГ) или словесный («ДД» месяц ГГГГ г.)
    3) Номер приказа: символ № + непустое значение
    4) Город: г. + название
    """
    scopes = ["шапка"]
    text = _get_scopes_text(target_doc, scopes)
    violations = []
    rule_index = 3
    rule_title = "Проверка реквизитов приказа"

    # 1) Юрлицо: ООО/ЗАО/АО/ПАО + текст в кавычках или без
    has_org = bool(re.search(
        r'(ООО|ЗАО|АО|ПАО)\s*[«"\']?.+',
        text
    ))
    if not has_org:
        violations.append({
            "rule_index": rule_index,
            "rule_title": rule_title,
            "Целевой документ": "отсутствует",
            "Различие": "Отсутствует наименование юрлица (ООО/ЗАО/АО/ПАО)"
        })

    # 2) Дата: числовой или словесный формат
    months = (
        r'январ[яь]|феврал[яь]|март[а]?|апрел[яь]|ма[йя]|'
        r'июн[яь]|июл[яь]|август[а]?|сентябр[яь]|октябр[яь]|'
        r'ноябр[яь]|декабр[яь]'
    )
    has_date = bool(re.search(
        r'\d{2}[.\s]+\d{2}[.\s]+\d{4}|'
        r'[«"\'"]?\d{1,2}[»"\'"]?\s*(?:' + months + r')',
        text, re.IGNORECASE
    ))
    # Плейсхолдер даты (__.__.202_) — не считаем
    if has_date and re.search(r'_+\._+\._+', text):
        has_date = False
    if not has_date:
        violations.append({
            "rule_index": rule_index,
            "rule_title": rule_title,
            "Целевой документ": "отсутствует",
            "Различие": "Отсутствует дата приказа"
        })

    # 3) Номер приказа: № + непустое значение
    number_match = re.search(r'№\s*(.+)', text)
    if number_match:
        after_sign = number_match.group(1).strip()
        cleaned = re.sub(r'[_\s]', '', after_sign)
        is_bn = bool(re.match(r'^б[/\\]?н$', cleaned, re.IGNORECASE))
        if not cleaned or is_bn:
            violations.append({
                "rule_index": rule_index,
                "rule_title": rule_title,
                "Целевой документ": f"№ {after_sign}",
                "Различие": "Номер приказа не заполнен"
            })
    else:
        violations.append({
            "rule_index": rule_index,
            "rule_title": rule_title,
            "Целевой документ": "отсутствует",
            "Различие": "Отсутствует символ № с номером приказа"
        })

    # 4) Город: г. + название
    has_city = bool(re.search(
        r'г\.\s*[А-ЯЁ][а-яё]+|город\s+[А-ЯЁ]|Москв[аеы]|Санкт-Петербург',
        text
    ))
    if not has_city:
        violations.append({
            "rule_index": rule_index,
            "rule_title": rule_title,
            "Целевой документ": "отсутствует",
            "Различие": "Отсутствует город подписания (например «г. Москва»)"
        })

    return violations


@register("prikaz_formirovanie_po", 4)
def check_stamp_formirovanie_po(target_doc, config):
    """Правило #4: должность + ФИО + М.П./ПЕЧАТЬ для Приказа о формировании ПО."""
    return check_signatory_with_stamp(target_doc, config, 4, "Проверка подписанта и М.П.")


# ============================================================================
# Регистрация для prikaz_ic (Приказ о создании ИЦ предприятия)
# ============================================================================

@register("prikaz_ic", 2)
def check_rekvizity_ic(target_doc, config):
    """
    Правило #2: реквизиты приказа (OCR-толерантный) для Приказа о создании ИЦ.

    Проверяет 5 обязательных реквизитов в шапке:
    1) Юрлицо: ООО/ЗАО/АО/ПАО (+ OCR: 000=ООО)
    2) ПРИКАЗ
    3) Дата
    4) Номер
    5) Город
    """
    scopes = ["шапка"]
    text = _get_scopes_text(target_doc, scopes)
    violations = []
    rule_index = 2
    rule_title = "Проверка реквизитов приказа"

    # 1) Юрлицо: ООО/000 (OCR-артефакт нулей вместо букв О) + текст
    has_org = bool(re.search(
        r'(ООО|000|OOO|ЗАО|АО|ПАО)\s*[«"\']?.+',
        text
    ))
    if not has_org:
        violations.append({
            "rule_index": rule_index,
            "rule_title": rule_title,
            "Целевой документ": "отсутствует",
            "Различие": "Отсутствует наименование юрлица (ООО/ЗАО/АО/ПАО)"
        })

    # 2) ПРИКАЗ
    has_prikaz = bool(re.search(
        r'П\s*Р\s*И\s*К\s*А\s*З|Приказ',
        text, re.IGNORECASE
    ))
    if not has_prikaz:
        violations.append({
            "rule_index": rule_index,
            "rule_title": rule_title,
            "Целевой документ": "отсутствует",
            "Различие": "Отсутствует слово «ПРИКАЗ»"
        })

    # 3) Дата: числовой или словесный формат
    months = (
        r'январ[яь]|феврал[яь]|март[а]?|апрел[яь]|ма[йя]|'
        r'июн[яь]|июл[яь]|август[а]?|сентябр[яь]|октябр[яь]|'
        r'ноябр[яь]|декабр[яь]'
    )
    has_date = bool(re.search(
        r'\d{2}[.\s]+\d{2}[.\s]+\d{4}|'
        r'[«"\'"]?\d{1,2}[»"\'"]?\s*(?:' + months + r')',
        text, re.IGNORECASE
    ))
    if has_date and re.search(r'_+\._+\._+', text):
        has_date = False
    if not has_date:
        violations.append({
            "rule_index": rule_index,
            "rule_title": rule_title,
            "Целевой документ": "отсутствует",
            "Различие": "Отсутствует дата приказа"
        })

    # 4) Номер приказа: № + непустое значение
    number_match = re.search(r'№\s*(.+)', text)
    if number_match:
        after_sign = number_match.group(1).strip()
        cleaned = re.sub(r'[_\s]', '', after_sign)
        is_bn = bool(re.match(r'^б[/\\]?н$', cleaned, re.IGNORECASE))
        if not cleaned or is_bn:
            violations.append({
                "rule_index": rule_index,
                "rule_title": rule_title,
                "Целевой документ": f"№ {after_sign}",
                "Различие": "Номер приказа не заполнен"
            })
    else:
        violations.append({
            "rule_index": rule_index,
            "rule_title": rule_title,
            "Целевой документ": "отсутствует",
            "Различие": "Отсутствует символ № с номером приказа"
        })

    # 5) Город: г. + название
    has_city = bool(re.search(
        r'г\.\s*[А-ЯЁ][а-яё]+|город\s+[А-ЯЁ]|Москв[аеы]|Санкт-Петербург',
        text
    ))
    if not has_city:
        violations.append({
            "rule_index": rule_index,
            "rule_title": rule_title,
            "Целевой документ": "отсутствует",
            "Различие": "Отсутствует город подписания (например «г. Москва»)"
        })

    return violations


@register("prikaz_ic", 4)
def check_signatory_ic(target_doc, config):
    """Правило #4: должность + ФИО подписанта для Приказа о создании ИЦ."""
    return check_signatory(target_doc, config, 4,
                           "Проверка подписанта.", scopes=["шапка"])


@register("prikaz_ic", 6)
def check_responsible_fio_ic(target_doc, config):
    """
    Правило #6: ФИО ответственных лиц в регламенте для Приказа о создании ИЦ.

    Переиспользует логику check_responsible_fio_ic_potoka с маппингом scope:
    prikaz_ic использует "приложение_2_к_приказу", а не "приложение_2_регламент".
    """
    # Адаптируем scope для совместимости с check_responsible_fio_ic_potoka
    adapted_doc = dict(target_doc)
    adapted_doc["приложение_2_регламент"] = target_doc.get("приложение_2_к_приказу", "")
    violations = check_responsible_fio_ic_potoka(adapted_doc, config)
    # Переписываем rule_index/title на актуальные для prikaz_ic
    for v in violations:
        v["rule_index"] = 6
        v["rule_title"] = "Проверка ФИО ответственных лиц в регламенте."
    return violations
