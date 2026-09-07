#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Согласованность реестра спецдвижков и конфигов типов.

Ловит ошибку «тип работает из командной строки, но молчит в веб-интерфейсе»:
каждый `engine` из doc_configs/*/config.json должен быть в реестре спецдвижков,
и реестр CLI (main.py) должен совпадать с реестром веб-бэкенда (src/api/server.py).
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
    """Реестры CLI и веб-бэкенда содержат один и тот же набор движков."""
    import main
    from src.api import server

    assert set(main.SPECIAL_ENGINE_RUNNERS) == set(server.SPECIAL_ENGINE_RUNNERS), (
        "реестры разошлись: только в CLI — "
        f"{set(main.SPECIAL_ENGINE_RUNNERS) - set(server.SPECIAL_ENGINE_RUNNERS)}, "
        f"только в веб — {set(server.SPECIAL_ENGINE_RUNNERS) - set(main.SPECIAL_ENGINE_RUNNERS)}"
    )


@pytest.mark.parametrize("doc_type,engine", sorted(_configured_engines().items()))
def test_every_configured_engine_is_registered(doc_type, engine):
    """Движок из config.json есть в реестре (иначе тип не запустится нигде)."""
    import main

    assert engine in main.SPECIAL_ENGINE_RUNNERS, f"{doc_type}: engine={engine!r} нет в реестре"


def test_every_registered_engine_is_used():
    """В реестре нет мёртвых записей — каждый движок используется хотя бы одним типом."""
    import main

    used = set(_configured_engines().values())
    dead = set(main.SPECIAL_ENGINE_RUNNERS) - used
    assert not dead, f"движки в реестре без конфига типа: {sorted(dead)}"
