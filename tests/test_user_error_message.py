#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Тесты F2: исключения аудита переводятся в человеческий текст, технический — отдельно.
"""
import json

import httpx
import openai
import pytest


def _api_error(cls, status):
    resp = httpx.Response(status, request=httpx.Request("POST", "http://x/v1/chat/completions"), json={"error": {"message": "m"}})
    return cls("m", response=resp, body=None)


@pytest.mark.parametrize("exc,fragment", [
    (openai.APITimeoutError(httpx.Request("POST", "http://x")), "не ответил за отведённое время"),
    (openai.APIConnectionError(request=httpx.Request("POST", "http://x")), "недоступен или не отвечает"),
    (_api_error(openai.NotFoundError, 404), "не найдена на сервере языковой модели"),
    (_api_error(openai.InternalServerError, 500), "вернул ошибку (код 500)"),
    (ConnectionError("Не удалось подключиться к VLM (http://127.0.0.1:1/v1/)"), "Сервис распознавания сканов недоступен"),
    (PermissionError(13, "Permission denied", "/x/logs_result"), "Ошибка записи на сервере"),
    (json.JSONDecodeError("Expecting value", "x", 0), "неожиданном формате"),
    (ValueError("Документ не содержит распознаваемого текста"), "Документ не содержит распознаваемого текста"),
    (RuntimeError("что-то"), "Не удалось выполнить проверку (RuntimeError)"),
])
def test_user_error_message(exc, fragment):
    from src.api.server import _user_error_message
    msg = _user_error_message(exc)
    assert fragment in msg
    # техника (классы, URL, пути) не должна попадать в пользовательский текст
    for bad in ("Traceback", "Errno", "http://", "APIConnectionError:"):
        assert bad not in msg
