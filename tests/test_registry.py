#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Согласованность реестра спецдвижков и конфигов типов.

Ловит ошибку «тип работает из командной строки, но молчит в веб-интерфейсе»:
каждый `engine` из doc_configs/*/config.json должен быть в реестре спецдвижков
(src/engines.py), а CLI (main.py) и веб-бэкенд (src/api/server.py) должны использовать
именно этот реестр, а не свои копии.
"""
import json

import pytest

from conftest import DOC_CONFIGS


def _configured_engines() -> dict:
    """{doc_type: engine} по всем config.json, у которых задано поле engine."""
    out = {}
    for cfg in sorted(DOC_CONFIGS.glob("*/config.json")):
        with open(cfg, encoding="utf-8") as f:
            data = json.load(f)
        if data.get("engine"):
            out[data["doc_type"]] = data["engine"]
    return out


def test_cli_and_web_registries_identical():
    """CLI и веб-бэкенд используют один объект реестра из src/engines.py."""
    import main
    from src import engines
    from src.api import server

    assert main.SPECIAL_ENGINE_RUNNERS is engines.SPECIAL_ENGINE_RUNNERS, "main.py держит свою копию реестра"
    assert server.SPECIAL_ENGINE_RUNNERS is engines.SPECIAL_ENGINE_RUNNERS, "server.py держит свою копию реестра"
    assert set(main.SPECIAL_ENGINE_RUNNERS) == set(server.SPECIAL_ENGINE_RUNNERS), (
        "реестры разошлись: только в CLI — "
        f"{set(main.SPECIAL_ENGINE_RUNNERS) - set(server.SPECIAL_ENGINE_RUNNERS)}, "
        f"только в веб — {set(server.SPECIAL_ENGINE_RUNNERS) - set(main.SPECIAL_ENGINE_RUNNERS)}"
    )


@pytest.mark.parametrize("doc_type,engine", sorted(_configured_engines().items()))
def test_every_configured_engine_is_registered(doc_type, engine):
    """Движок из config.json есть в реестре (иначе тип не запустится нигде)."""
    from src.engines import SPECIAL_ENGINE_RUNNERS

    assert engine in SPECIAL_ENGINE_RUNNERS, f"{doc_type}: engine={engine!r} нет в реестре"


def test_every_registered_engine_is_used():
    """В реестре нет мёртвых записей — каждый движок используется хотя бы одним типом."""
    from src.engines import SPECIAL_ENGINE_RUNNERS

    used = set(_configured_engines().values())
    dead = set(SPECIAL_ENGINE_RUNNERS) - used
    assert not dead, f"движки в реестре без конфига типа: {sorted(dead)}"
