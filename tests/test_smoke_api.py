#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Smoke-тесты HTTP API (без запуска сервера — проверяем импорт и базовые маршруты).

Запускаем FastAPI через TestClient — быстро, без сети.
"""
import pytest

pytestmark = pytest.mark.skipif(
    True,  # подменится ниже если TestClient доступен
    reason="httpx/TestClient не установлен",
)


def _testclient_available():
    try:
        from fastapi.testclient import TestClient  # noqa: F401
        return True
    except ImportError:
        return False


if _testclient_available():
    pytestmark = pytest.mark.skipif(False, reason="")


def test_api_import():
    """api_server.py импортируется без падений."""
    import api_server
    assert hasattr(api_server, "app"), "FastAPI app отсутствует"


def test_api_health():
    """GET /api/health отвечает (status=ok или error — без 500)."""
    from fastapi.testclient import TestClient
    import api_server

    client = TestClient(api_server.app)
    response = client.get("/api/health")
    assert response.status_code == 200, f"Health endpoint упал: {response.status_code}"
    data = response.json()
    assert "status" in data


def test_list_types_requires_auth():
    """GET /api/types требует basic auth — без auth 401."""
    from fastapi.testclient import TestClient
    import api_server

    client = TestClient(api_server.app)
    response = client.get("/api/types")
    assert response.status_code == 401, "Защищённый эндпоинт должен отвергать unauth"


def test_list_types_with_auth():
    """POST /api/login → GET /api/types с cookie — 200, возвращает типы."""
    import os
    from fastapi.testclient import TestClient
    import api_server

    client = TestClient(api_server.app)
    login = os.getenv("AUDIT_LOGIN", "admin")
    password = os.getenv("AUDIT_PASSWORD", "mos186124kva")

    # Логин
    login_resp = client.post("/api/login", json={"login": login, "password": password})
    assert login_resp.status_code == 200, f"Логин упал: {login_resp.status_code}"

    # Запрос защищённого эндпоинта с cookie из клиента (TestClient хранит автоматически)
    response = client.get("/api/types")
    assert response.status_code == 200, f"С cookie должно быть 200, got {response.status_code}"
    data = response.json()
    assert isinstance(data, (list, dict))
