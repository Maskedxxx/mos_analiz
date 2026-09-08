#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Тесты F6: /api/health отвечает 503 при degraded (недоступен LLM или layout) и включает layout-детектор.
Адреса подменяются копиями конфигов (pydantic-модели frozen — правим через model_copy).
"""
from fastapi.testclient import TestClient


def _health(monkeypatch, *, llm_url=None, layout_url=None):
    import main
    import src.api.server as srv
    if llm_url:
        monkeypatch.setattr(srv, "LLM_CONFIG", srv.LLM_CONFIG.model_copy(update={"base_url": llm_url}))
    if layout_url:
        pdf = srv.PARSERS_CONFIG.pdf
        layout = pdf.layout.model_copy(update={"base_url": layout_url})
        monkeypatch.setattr(srv, "PARSERS_CONFIG", srv.PARSERS_CONFIG.model_copy(update={"pdf": pdf.model_copy(update={"layout": layout})}))
    return TestClient(main.app).get("/api/health")


def test_llm_down_gives_503(monkeypatch):
    r = _health(monkeypatch, llm_url="http://127.0.0.1:1/v1/")
    assert r.status_code == 503
    body = r.json()
    assert body["status"] == "degraded"
    assert body["services"]["llm"]["status"] == "error"


def test_layout_down_gives_503_and_is_listed(monkeypatch):
    r = _health(monkeypatch, layout_url="http://127.0.0.1:2")
    assert r.status_code == 503
    layout = r.json()["services"]["layout"]
    assert layout["status"] == "error"
    assert layout["url"] == "http://127.0.0.1:2/health"


def test_layout_present_in_health(monkeypatch):
    r = _health(monkeypatch)
    assert "layout" in r.json()["services"]
    assert r.status_code in (200, 503)
