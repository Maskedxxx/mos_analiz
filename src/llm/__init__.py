# START_MODULE_CONTRACT
# PURPOSE: Пакет LLM-слоя. На данный момент содержит клиента (`client.py`). В следующих подшагах (step 2.3/2.4) добавятся `context_builder.py` и `multi_rule.py`.
# INPUTS: Импорт-время — подмодули пакета.
# OUTPUTS: Публичные функции для вызывающего кода: `call_llm`, `parse_json_response`, `resolve_runtime_llm_model`.
# KEYWORDS: package, llm, public-api.
# LINKS: src/llm/client.py, main.py (валидаторы + multi_rule).
# RATIONALE: Снаружи пакет воспринимается как «LLM-сервис». Внутренности (OpenAI SDK, JSON-санитайзер) — package-private.
# END_MODULE_CONTRACT

from __future__ import annotations

# START_REEXPORTS
from src.llm.client import (
    call_llm,
    parse_json_response,
    resolve_runtime_llm_model,
    sanitize_json_string,
)

__all__ = [
    "call_llm",
    "parse_json_response",
    "resolve_runtime_llm_model",
    "sanitize_json_string",
]
# END_REEXPORTS
