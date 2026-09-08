#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Тесты F16 (нечитаемый ответ модели → повтор со сменой seed, после 5 попыток — человеческий текст)
и F10 (документ больше контекста модели → понятный отказ до вызова модели).
Клиент модели подменяется: параметры вызова не трогаются, считается число вызовов и seed.
"""
import json

import pytest


class _Msg:
    def __init__(self, c): self.content = c
class _Choice:
    def __init__(self, c): self.message = _Msg(c)
class _Usage:
    prompt_tokens = 1; completion_tokens = 1; total_tokens = 2
class _Resp:
    def __init__(self, c): self.choices = [_Choice(c)]; self.usage = _Usage()


class _Completions:
    """Отдаёт по очереди заготовленные ответы; фиксирует seed каждого вызова."""
    def __init__(self, answers): self.answers = list(answers); self.calls = []
    def create(self, **kw):
        self.calls.append(kw.get("seed"))
        return _Resp(self.answers[min(len(self.calls) - 1, len(self.answers) - 1)])


class _Client:
    def __init__(self, answers):
        self.chat = type("Chat", (), {})(); self.chat.completions = _Completions(answers)


RULES = [{"index": 1, "title": "П1", "check": "c", "target_sections": []},
         {"index": 2, "title": "П2", "check": "c", "target_sections": []}]
GOOD = json.dumps([{"rule_index": 1, "reasoning": "r", "verdict": {"status": "ok"}},
                   {"rule_index": 2, "reasoning": "r", "verdict": {"status": "ok"}}], ensure_ascii=False)


def _run(monkeypatch, answers):
    import src.llm.multi_rule as mr
    client = _Client(answers)
    monkeypatch.setattr(mr, "make_llm_client", lambda base_url=None: client)
    res = mr.run_multi_rule_audit(parsed={"filename": "d", "path": "d", "raw_text": "текст"}, sections={}, rules=RULES,
                                  include_scopes=None, filename="d", llm_base_url="http://stub/v1/", llm_model="m")
    return res, client.chat.completions.calls


def test_unparseable_then_good_is_retried(monkeypatch):
    res, seeds = _run(monkeypatch, ["Извините, не могу.", "", '[{"rule_index": 1,', GOOD])
    assert len(seeds) == 4 and seeds == [42, 43, 44, 45]
    assert res["checked_count"] == 2


def test_all_unparseable_gives_human_error(monkeypatch):
    with pytest.raises(ValueError, match="не вернула результат в ожидаемом формате после 5 попыток"):
        _run(monkeypatch, ["не JSON"])


def test_context_guard():
    from src.llm.multi_rule import check_prompt_fits_context
    check_prompt_fits_context(3000 * 3, 262144)
    with pytest.raises(ValueError, match="слишком большой для проверки"):
        check_prompt_fits_context(262144 * 3 + 100, 262144)


def test_context_guard_wired(monkeypatch):
    import src.llm.multi_rule as mr
    monkeypatch.setattr(mr.LLM_CONFIG, "context_tokens", 10, raising=False) if not getattr(type(mr.LLM_CONFIG), "model_config", {}).get("frozen") else None
    # LLM_CONFIG заморожен — подменяем через объект-двойник
    fake = type("Cfg", (), {"context_tokens": 10, "default_model": "m"})()
    monkeypatch.setattr(mr, "LLM_CONFIG", fake)
    client = _Client([GOOD])
    monkeypatch.setattr(mr, "make_llm_client", lambda base_url=None: client)
    with pytest.raises(ValueError, match="слишком большой"):
        mr.run_multi_rule_audit(parsed={"filename": "d", "path": "d", "raw_text": "x" * 1000}, sections={}, rules=RULES,
                                include_scopes=None, filename="d", llm_base_url="u", llm_model="m")
    assert client.chat.completions.calls == []  # модель не вызывалась


def test_context_error_mapping():
    import httpx, openai
    from src.api.server import _user_error_message
    resp = httpx.Response(400, request=httpx.Request("POST", "http://x"), json={"error": {"message": "request (374868 tokens) exceeds the available context size (262144 tokens)"}})
    exc = openai.BadRequestError("request (374868 tokens) exceeds the available context size", response=resp, body=None)
    assert "слишком большой" in _user_error_message(exc)
