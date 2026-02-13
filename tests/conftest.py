#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Общие фикстуры для тестов аудит-движка.

Содержит пути к тестовым файлам и вспомогательные функции.
"""
import sys
from pathlib import Path

import pytest

# Корень проекта
PROJECT_ROOT = Path(__file__).parent.parent
# Добавляем audit_engine в sys.path для импортов
sys.path.insert(0, str(PROJECT_ROOT / "audit_engine"))
sys.path.insert(0, str(PROJECT_ROOT))

TEST_DOCS = PROJECT_ROOT / "test_docs"
DOC_CONFIGS = PROJECT_ROOT / "doc_configs"


# === Пути к тестовым файлам kpsc ===
KPSC_FILES = {
    "biznes_otel": TEST_DOCS / "kpsc" / "biznes_otel" / "КПСЦ_ТС_ООО_Бизнес_Отель_v3.xlsx",
    "mapper": TEST_DOCS / "kpsc" / "mapper" / "КПСЦ Текущее состояние ООО МАППЕР версия 2.xlsx",
    "rotosnab": TEST_DOCS / "kpsc" / "rotosnab" / "КПСЦ_и_Спагетти_ТС_ООО_Ротоснаб_итог.xlsx",
    "ruslet": TEST_DOCS / "kpsc" / "ruslet" / "КПСЦ и Спагетти ТС_ООО РУСЛЕТ_v.04.1.xlsx",
    "sodex": TEST_DOCS / "kpsc" / "sodex" / "КПСЦ Текущее состояние ООО СОДЕКС версия 4 (1).xlsx",
}


# === Пути к тестовым файлам prikaz_comp_ppu ===
PRIKAZ_COMP_PPU_FILES = {
    "biznes_otel": TEST_DOCS / "prikaz_comp_ppu" / "biznes_otel" / "3_ПРИКАЗ_о_конкурсах_ППУ_ООО_Бизнес_отель.docx",
    "mapper": TEST_DOCS / "prikaz_comp_ppu" / "mapper" / "ПРИКАЗ о конкурсах ППУ Маппер.docx",
    "rotosnab": TEST_DOCS / "prikaz_comp_ppu" / "rotosnab" / "ПРИКАЗ о конкурсах ППУ.docx",
    "ruslet": TEST_DOCS / "prikaz_comp_ppu" / "ruslet" / "Приказ о проведении конкурсов ППУ_от 14.10.2025.docx",
    "sodex": TEST_DOCS / "prikaz_comp_ppu" / "sodex" / "ПРИКАЗ о конкурсах ППУ ООО Содекс.docx",
}


@pytest.fixture
def kpsc_files():
    """Словарь путей к тестовым XLSX файлам kpsc."""
    return KPSC_FILES


@pytest.fixture
def prikaz_comp_ppu_files():
    """Словарь путей к тестовым DOCX файлам prikaz_comp_ppu."""
    return PRIKAZ_COMP_PPU_FILES
