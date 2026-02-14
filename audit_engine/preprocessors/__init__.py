#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Препроцессоры текста — нормализация перед отправкой в LLM.

Регистрируются через декоратор @register_preprocessor(doc_type, scope).
Движок вызывает зарегистрированные препроцессоры при build_context_for_rule().
"""

from .registry import register_preprocessor, get_preprocessors

# Импортируем модули с препроцессорами для автоматической регистрации
from . import prikaz_ic
from . import cheklist
from . import presentation_eu
from . import iter8_normalize
from . import prikaz_ic_potoka
from . import prikaz_vyhod
from . import prikaz_comp_ppu
from . import prikaz_ppu
from . import polozhenie_normalize
from . import polozhenie_po
