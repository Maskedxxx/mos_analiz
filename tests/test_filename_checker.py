#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Тесты для check_filename_universal — проверка имени файла.

Покрывает два режима:
1. filename_keywords — список ключевых слов (новый)
2. filename_pattern — подстрока (обратная совместимость)
"""
from dataclasses import dataclass
from typing import List, Optional

from main import check_filename_universal


@dataclass
class MockConfig:
    """Мок для AuditConfig — только поля нужные для filename checker."""
    filename_pattern: str = ""
    filename_keywords: Optional[List[str]] = None


# === Тесты filename_keywords (новый режим) ===

class TestFilenameKeywords:
    """Тесты для режима filename_keywords."""

    def test_keywords_all_present(self):
        """Все ключевые слова есть в имени → PASS."""
        config = MockConfig(filename_keywords=["приказ", "ппу"])
        target = {"filename": "ПРИКАЗ о конкурсах ППУ ООО Содекс.docx"}
        result = check_filename_universal(target, config)
        assert result == [], f"Ожидался PASS, но получен FAIL: {result}"

    def test_keywords_case_insensitive(self):
        """Проверка регистронезависимости."""
        config = MockConfig(filename_keywords=["приказ", "ппу"])
        target = {"filename": "3_ПРИКАЗ_о_конкурсах_ППУ_ООО_Бизнес_отель.docx"}
        result = check_filename_universal(target, config)
        assert result == []

    def test_keywords_missing_one(self):
        """Одно ключевое слово отсутствует → FAIL."""
        config = MockConfig(filename_keywords=["приказ", "ппу"])
        target = {"filename": "Положение о ППУ.docx"}
        result = check_filename_universal(target, config)
        assert len(result) == 1
        assert "приказ" in result[0]["Различие"].lower()

    def test_keywords_missing_all(self):
        """Все ключевые слова отсутствуют → FAIL."""
        config = MockConfig(filename_keywords=["положение", "ппу"])
        target = {"filename": "Документ.docx"}
        result = check_filename_universal(target, config)
        assert len(result) == 1

    def test_keywords_single_word(self):
        """Одно ключевое слово — достаточно."""
        config = MockConfig(filename_keywords=["ппу"])
        target = {"filename": "ПРД-РТ-11-002_ППУ.docx"}
        result = check_filename_universal(target, config)
        assert result == []

    def test_keywords_with_underscores(self):
        """Подчёркивания в имени файла не мешают."""
        config = MockConfig(filename_keywords=["положение", "ппу"])
        target = {"filename": "2_Положение_о_ППУ_ООО_Бизнес_отель.docx"}
        result = check_filename_universal(target, config)
        assert result == []

    # Реальные файлы из тестового прогона

    def test_prikaz_comp_ppu_all_companies(self):
        """prikaz_comp_ppu: все 5 компаний должны пройти с keywords."""
        config = MockConfig(filename_keywords=["приказ", "ппу"])
        filenames = [
            "3_ПРИКАЗ_о_конкурсах_ППУ_ООО_Бизнес_отель.docx",
            "ПРИКАЗ о конкурсах ППУ Маппер.docx",
            "ПРИКАЗ о конкурсах ППУ.docx",
            "Приказ о проведении конкурсов ППУ_от 14.10.2025.docx",
            "ПРИКАЗ о конкурсах ППУ ООО Содекс.docx",
        ]
        for fname in filenames:
            target = {"filename": fname}
            result = check_filename_universal(target, config)
            assert result == [], f"FAIL для файла: {fname}"

    def test_prikaz_ppu_all_companies(self):
        """prikaz_ppu: все 3 компании должны пройти."""
        config = MockConfig(filename_keywords=["приказ", "ппу"])
        filenames = [
            "1. ПРИКАЗ о ППУ ООО Бизнес-отель.docx",
            "ПРИКАЗ о ППУ.docx",
            "ПРИКАЗ о ППУ ООО Содекс.docx",
        ]
        for fname in filenames:
            target = {"filename": fname}
            result = check_filename_universal(target, config)
            assert result == [], f"FAIL для файла: {fname}"

    def test_polozhenie_ppu_all_companies(self):
        """polozhenie_ppu: все 4 компании должны пройти с ["ппу"]."""
        config = MockConfig(filename_keywords=["ппу"])
        filenames = [
            "2_Положение_о_ППУ_ООО_Бизнес_отель.docx",
            "Положение о ППУ.docx",
            "ПРД-РТ-11-002_ППУ.docx",  # ruslet — процедура, нет слова "положение"
            "Положение о ППУ ООО Содекс.docx",
        ]
        for fname in filenames:
            target = {"filename": fname}
            result = check_filename_universal(target, config)
            assert result == [], f"FAIL для файла: {fname}"

    def test_polozhenie_comp_ppu_all_companies(self):
        """polozhenie_comp_ppu: все 3 компании должны пройти."""
        config = MockConfig(filename_keywords=["положение", "ппу"])
        filenames = [
            "4_Положение_о_конкурсах_проектов_и_ППУ_ООО_Бизнес_отель.docx",
            "Положение_о_конкурсах_проектов_и_ППУ.docx",
            "Положение о конкурсах проектов и ППУ ООО СОДЕКС.docx",
        ]
        for fname in filenames:
            target = {"filename": fname}
            result = check_filename_universal(target, config)
            assert result == [], f"FAIL для файла: {fname}"


# === Тесты filename_pattern (обратная совместимость) ===

class TestFilenamePattern:
    """Тесты для старого режима filename_pattern."""

    def test_pattern_present(self):
        """Паттерн есть в имени → PASS."""
        config = MockConfig(filename_pattern="1.5 Приказ о создании ИЦ с прилож")
        target = {"filename": "1.5 Приказ о создании ИЦ с прилож.docx"}
        result = check_filename_universal(target, config)
        assert result == []

    def test_pattern_missing(self):
        """Паттерн отсутствует → FAIL."""
        config = MockConfig(filename_pattern="1.5 Приказ о создании ИЦ с прилож")
        target = {"filename": "Приказ_о_создании_ИЦ.docx"}
        result = check_filename_universal(target, config)
        assert len(result) == 1

    def test_pattern_empty(self):
        """Пустой паттерн → PASS (нет проверки)."""
        config = MockConfig(filename_pattern="")
        target = {"filename": "anything.docx"}
        result = check_filename_universal(target, config)
        assert result == []


# === Тесты приоритета: keywords > pattern ===

class TestKeywordsPriority:
    """Keywords имеют приоритет над pattern."""

    def test_keywords_override_pattern(self):
        """Если есть keywords — pattern игнорируется."""
        config = MockConfig(
            filename_pattern="order_comp_ppu",  # не совпадёт
            filename_keywords=["приказ", "ппу"]  # совпадёт
        )
        target = {"filename": "ПРИКАЗ о конкурсах ППУ.docx"}
        result = check_filename_universal(target, config)
        assert result == [], "keywords должны иметь приоритет над pattern"

    def test_no_keywords_falls_back_to_pattern(self):
        """Без keywords — используется pattern."""
        config = MockConfig(
            filename_pattern="1.5 Приказ",
            filename_keywords=None
        )
        target = {"filename": "1.5 Приказ о создании ИЦ.docx"}
        result = check_filename_universal(target, config)
        assert result == []
