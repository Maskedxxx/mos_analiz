# START_MODULE_CONTRACT
# PURPOSE: Конфиг LLM-сервиса. Pydantic-модель с дефолтными значениями и `description` для каждого поля. Заменяет хардкоженные env-var fallback-ы (LLM_BASE_URL, LLM_MODEL) и локальные дефолты в main.py.
# INPUTS: —
# OUTPUTS: `LlmConfig` и `LLM_CONFIG` (singleton на весь рантайм).
# KEYWORDS: config, pydantic, llm, spark, vllm.
# LINKS: src/llm/client.py (call_llm, resolve_runtime_llm_model), main.py (AuditConfig + kpsc-валидаторы).
# RATIONALE: LLM-инфраструктура — один сервис на весь проект (Spark-vLLM на 172.16.10.35:11437). Нет причин держать URL/модель/api_key разбросанными по env-vars и локальным литералам. Per-doc_type оверрайды (model/temperature/seed) остаются в `doc_configs/<тип>/config.json` через `AuditConfig`.
# END_MODULE_CONTRACT

from __future__ import annotations

# START_IMPORTS
from typing import Optional

from pydantic import BaseModel, ConfigDict, Field
# END_IMPORTS


# START_LLM_CONFIG
class LlmConfig(BaseModel):
    """Параметры LLM-сервиса: URL, дефолтная модель, лимиты, reasoning, seed."""

    model_config = ConfigDict(frozen=True, extra="forbid")

    base_url: str = Field(
        default="http://172.16.10.35:11437/v1/",
        description=(
            "URL LLM-сервиса (Spark vLLM). Используется `call_llm` в "
            "`src/llm/client.py` и напрямую валидаторами KPSC/kartochka. "
            "Per-doc_type оверрайд возможен через `AuditConfig.llm_base_url`."
        ),
    )
    api_key: str = Field(
        default="none",
        description=(
            "API-ключ. Для vLLM обычно любое непустое значение (не проверяется)."
        ),
    )
    default_model: str = Field(
        default="Qwen3.5-35B-A3B",
        description=(
            "Имя модели по умолчанию в Spark-vLLM. Используется как fallback в "
            "`resolve_runtime_llm_model`: если doc_type указывает стейл-идентификатор "
            "(`gpt-4.1-mini`, `openai/gpt-oss-120b`) — он подменяется на это значение."
        ),
    )
    default_max_tokens: int = Field(
        default=4096,
        description=(
            "Лимит токенов ответа по умолчанию. Используется, если `AuditConfig.llm_max_tokens` "
            "не задан в doc_type-конфиге."
        ),
    )
    default_temperature: float = Field(
        default=0.0,
        description=(
            "Температура генерации по умолчанию. Для аудита документов всегда 0 — "
            "нужен детерминизм, а не креативность."
        ),
    )
    default_reasoning_effort: Optional[str] = Field(
        default=None,
        description=(
            "Для gpt-oss-моделей на vLLM: `low`/`medium`/`high`. Прокидывается через "
            "`extra_body={'reasoning_effort': ...}` в chat.completions.\n"
            "\n"
            "**ВАЖНО:** по умолчанию `None` — т.е. параметр НЕ отправляется, и это "
            "сознательно. Наш рантайм-LLM (`Qwen3.5-35B-A3B` на Spark через "
            "`resolve_runtime_llm_model`) reasoning_effort НЕ поддерживает; включение "
            "ломает ответы (исторически выявлено). Связанная защита: "
            "`src/llm/client.py:call_llm` ВСЕГДА добавляет в `extra_body` "
            "`chat_template_kwargs={'enable_thinking': False}` чтобы Qwen не уходил в "
            "thinking-mode (он ломает строгий JSON, обёртывает в markdown).\n"
            "\n"
            "Doc_types, которым это поле реально нужно (gpt-oss совместимые пайплайны), "
            "выставляют его явно в `doc_configs/<тип>/config.json::reasoning_effort`."
        ),
    )
    default_seed: Optional[int] = Field(
        default=42,
        description=(
            "Фиксированный seed для воспроизводимости. Критически важен для MoE-моделей "
            "(например, gpt-oss) — без seed ответы сильно скачут. `None` → не фиксировать."
        ),
    )
# END_LLM_CONFIG


# START_SINGLETON
# PURPOSE: Единственный экземпляр LLM-конфига на весь рантайм.
# INPUTS: —
# OUTPUTS: `LLM_CONFIG`.
# KEYWORDS: singleton.
LLM_CONFIG = LlmConfig()
# END_SINGLETON
