# START_MODULE_CONTRACT
# PURPOSE: Конфиг LLM-сервиса. Pydantic-settings-модель: дефолты в коде, поверх — переменные окружения (LLM_BASE_URL, LLM_MODEL, LLM_API_KEY, LLM_MAX_TOKENS, LLM_SEED) и файл `.env` в корне проекта. Каждое поле с `description`.
# INPUTS: —
# OUTPUTS: `LlmConfig` и `LLM_CONFIG` (singleton на весь рантайм).
# KEYWORDS: config, pydantic-settings, env, llm, openai-compatible.
# LINKS: src/llm/client.py (call_llm, resolve_runtime_llm_model), main.py (AuditConfig + kpsc-валидаторы).
# RATIONALE: Одна точка правды об LLM-сервисе. Значения по умолчанию — рабочий стенд; при переносе на другой стенд достаточно задать переменные окружения (или `.env`), код и конфиги типов не трогаются. Приоритет: переменная окружения > `.env` > дефолт в коде. Per-doc_type оверрайды (model/temperature/seed/llm_base_url) по-прежнему возможны в `doc_configs/<тип>/config.json` через `AuditConfig`, но по умолчанию не нужны.
# END_MODULE_CONTRACT

from __future__ import annotations

# START_IMPORTS
from pathlib import Path
from typing import Optional

from pydantic import AliasChoices, Field
from pydantic_settings import BaseSettings, SettingsConfigDict

# Корень проекта: config/llm.py → parents[1]. Отсюда читается `.env`, если он есть.
_PROJECT_ROOT = Path(__file__).resolve().parents[1]
# END_IMPORTS


# START_LLM_CONFIG
class LlmConfig(BaseSettings):
    """Параметры LLM-сервиса: URL, дефолтная модель, лимиты, reasoning, seed.

    Источники значений (по убыванию приоритета): переменные окружения с префиксом `LLM_`,
    файл `.env` в корне проекта, дефолты ниже. Имена переменных — в описании каждого поля.
    """

    model_config = SettingsConfigDict(
        frozen=True,
        extra="ignore",
        env_prefix="LLM_",
        env_file=str(_PROJECT_ROOT / ".env"),
        env_file_encoding="utf-8",
    )

    base_url: str = Field(
        default="http://127.0.0.1:11437/v1/",
        description=(
            "URL OpenAI-совместимого LLM-сервиса (env: LLM_BASE_URL). Используется `call_llm` в "
            "`src/llm/client.py` и напрямую валидаторами KPSC/kartochka. "
            "Per-doc_type оверрайд возможен через `AuditConfig.llm_base_url`."
        ),
    )
    api_key: str = Field(
        default="none",
        description=(
            "API-ключ (env: LLM_API_KEY). Для vLLM/llama.cpp обычно любое непустое значение (не проверяется)."
        ),
    )
    default_model: str = Field(
        default="Qwen3.6-35B-A3B",
        validation_alias=AliasChoices("LLM_MODEL", "LLM_DEFAULT_MODEL"),
        description=(
            "Имя модели по умолчанию (env: LLM_MODEL) — как её зарегистрировал LLM-сервис "
            "(`GET <LLM_BASE_URL>/models`). Используется как fallback в "
            "`resolve_runtime_llm_model`: если doc_type указывает стейл-идентификатор "
            "(`gpt-4.1-mini`, `openai/gpt-oss-120b`) — он подменяется на это значение."
        ),
    )
    default_max_tokens: int = Field(
        default=4096,
        validation_alias=AliasChoices("LLM_MAX_TOKENS", "LLM_DEFAULT_MAX_TOKENS"),
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
            "сознательно. Рантайм-модель Qwen (через `resolve_runtime_llm_model`) "
            "reasoning_effort НЕ поддерживает; включение "
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
        validation_alias=AliasChoices("LLM_SEED", "LLM_DEFAULT_SEED"),
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
