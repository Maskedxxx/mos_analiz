#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Интеграционные smoke-тесты полного pipeline.

Запускают engine.run() end-to-end на одном файле каждого типа.
Медленные (LLM-вызовы) — запускать через pytest -m integration.

Требования: Layout API + Paddle OCR + vLLM доступны на 172.16.10.35.
"""
import os
import tempfile
import urllib.request
from pathlib import Path

import pytest

from conftest import LOCAL_SMOKE_FIXTURES, PROJECT_ROOT


# По одному быстрому docx на каждый movement-engine type — для ускорения
# Берём простые документы без тяжёлых приложений
PIPELINE_FIXTURES = {
    "akt_nachala": str(LOCAL_SMOKE_FIXTURES["akt_nachala"].relative_to(PROJECT_ROOT)),
    "prikaz_ppu": str(LOCAL_SMOKE_FIXTURES["prikaz_ppu"].relative_to(PROJECT_ROOT)),
    "cheklist_eu": str(LOCAL_SMOKE_FIXTURES["cheklist_eu"].relative_to(PROJECT_ROOT)),
}


def _services_available() -> bool:
    """Проверяет что LLM, OCR и Layout API доступны."""
    for url in [
        "http://172.16.10.35:11437/v1/models",
        "http://172.16.10.35:11438/v1/models",
        "http://172.16.10.35:11439/health",
    ]:
        try:
            with urllib.request.urlopen(url, timeout=3) as r:
                if r.status != 200:
                    return False
        except Exception:
            return False
    return True


@pytest.mark.integration
@pytest.mark.parametrize("doc_type,rel_path", sorted(PIPELINE_FIXTURES.items()))
def test_pipeline_end_to_end(doc_type, rel_path):
    """Полный прогон engine.run() — не крашится, возвращает result с violations."""
    if not _services_available():
        pytest.skip("ML-сервисы на 172.16.10.35 недоступны")

    file_path = PROJECT_ROOT / rel_path
    if not file_path.exists():
        pytest.skip(f"Фикстура не найдена: {rel_path}")

    os.environ["NO_PROXY"] = "172.16.10.35,localhost,127.0.0.1"

    from main import AuditEngine

    with tempfile.TemporaryDirectory() as tmpdir:
        engine = AuditEngine(doc_type)
        result = engine.run(str(file_path), session_dir=tmpdir)

        # Не крашнулся — главный smoke-критерий
        assert result is not None, "engine.run вернул None"

        # Структура результата
        assert hasattr(result, "violations"), "У result нет поля violations"
        assert isinstance(result.violations, list), "violations должен быть list"

        # Артефакты сессии созданы
        session_dir = Path(result.session_dir) if hasattr(result, "session_dir") else Path(tmpdir)
        assert (session_dir / "pipeline.log").exists(), "Нет pipeline.log"
        assert (session_dir / "parsed_docs").exists(), "Нет parsed_docs/"
