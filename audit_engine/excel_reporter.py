#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Экспорт результатов аудита в Excel.

Формирует отчёт с колонками:
- rule_index — номер правила
- rule_title — название правила
- Целевой документ — что фактически в документе
- Различие — что должно быть или в чём проблема
"""

from pathlib import Path
from typing import Any, Dict, List

import pandas as pd


def save_to_excel(violations: List[Dict[str, Any]], output_path: str) -> None:
    """
    Сохраняет результаты аудита в Excel.

    Args:
        violations: список нарушений
        output_path: путь для сохранения .xlsx
    """
    if not violations:
        # Пустой отчёт со стандартными колонками
        df = pd.DataFrame(columns=[
            "rule_index",
            "rule_title",
            "Целевой документ",
            "Различие"
        ])
    else:
        df = pd.DataFrame(violations)
        # Упорядочиваем колонки: стандартные первые, остальные после
        columns_order = ["rule_index", "rule_title", "Целевой документ", "Различие"]
        existing_cols = [c for c in columns_order if c in df.columns]
        extra_cols = [c for c in df.columns if c not in columns_order]
        df = df[existing_cols + extra_cols]

    Path(output_path).parent.mkdir(parents=True, exist_ok=True)
    df.to_excel(output_path, index=False, engine='openpyxl')
