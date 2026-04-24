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
from src.llm.multi_rule import (
    RULE_BLOCK_TEMPLATE,
    SECTION_BLOCK_TEMPLATE,
    SYSTEM_PROMPT,
    USER_PROMPT_TEMPLATE,
    build_user_prompt as multi_rule_build_user_prompt,
    load_methodology_config,
    load_multi_rule_config,
    run_multi_rule_audit,
)

__all__ = [
    # client
    "call_llm",
    "parse_json_response",
    "resolve_runtime_llm_model",
    "sanitize_json_string",
    # multi_rule
    "SYSTEM_PROMPT",
    "USER_PROMPT_TEMPLATE",
    "SECTION_BLOCK_TEMPLATE",
    "RULE_BLOCK_TEMPLATE",
    "multi_rule_build_user_prompt",
    "run_multi_rule_audit",
    "load_multi_rule_config",
    "load_methodology_config",
]
# END_REEXPORTS
