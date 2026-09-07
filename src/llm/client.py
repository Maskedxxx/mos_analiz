# START_MODULE_CONTRACT
# PURPOSE: Единственное место создания клиентов к OpenAI-совместимым сервисам (`make_llm_client` / `make_async_llm_client`: base_url, api_key, таймаут — из config/llm.py) и универсальный вызов `call_llm`. Плюс утилиты: парсинг JSON-ответа (с обработкой markdown-блоков и битых переносов строк) и резолвинг стейл-идентификаторов облачных моделей в локальное имя.
# INPUTS: Список messages (OpenAI-совместимый), параметры модели (name, base_url, temperature, max_tokens, reasoning_effort, seed).
# OUTPUTS: `call_llm` → текстовый ответ; `parse_json_response` → список нарушений. Errors: JSONDecodeError → пустой список + warning в stderr.
# KEYWORDS: llm, openai, json-parse, sanitize, reasoning-effort, vllm.
# LINKS: src/llm/multi_rule.py (будет на 2.4), main.py (kpsc/kartochka validators), config/parsers.py (для будущего LlmConfig).
# RATIONALE: Точка подключения модели должна быть одна — чтобы смена сервиса (адрес, ключ, таймаут) не требовала правок по модулям. Клиент создаётся только здесь; сами вызовы `chat.completions.create` с особыми параметрами остаются в модулях, где они нужны: `src/llm/multi_rule.py` (Qwen-специфичные параметры и ретраи), `src/doc_type_validators/crosscheck.py`, `src/api/server.py` (автор правил), `src/format_parsers/pdf/_clients.py` (VLM). Всё остальное ходит через `call_llm`.
# END_MODULE_CONTRACT

from __future__ import annotations

# START_IMPORTS
import json
import os
import sys
from typing import Any, Dict, List, Optional

from openai import AsyncOpenAI, OpenAI

from config.llm import LLM_CONFIG
# END_IMPORTS


# START_OPENAI_DEFAULTS
# PURPOSE: Защита от висящих LLM-запросов и принудительное `enable_thinking=False`
# для Qwen-thinking моделей на Spark-vLLM. Эти параметры применяются ВСЕГДА ко
# всем `call_llm`-вызовам (раньше делалось через monkey-patch в main.py — теперь
# inline в клиенте).
# - OPENAI_TIMEOUT_SEC: дефолтный таймаут запроса (сек), env-override.
# - DISABLE_THINKING_EXTRA_BODY: всегда добавляется в `extra_body` чтобы Qwen
#   не уходил в thinking-mode (он ломает строгий JSON, обёртывает в markdown).
OPENAI_TIMEOUT_SEC = float(os.environ.get("OPENAI_TIMEOUT_SEC", "300"))
DISABLE_THINKING_EXTRA_BODY: Dict[str, Any] = {"chat_template_kwargs": {"enable_thinking": False}}
# END_OPENAI_DEFAULTS


# START_CLIENT_FACTORY
# PURPOSE: Единственное место, где создаются клиенты к OpenAI-совместимым сервисам (LLM, VLM/OCR).
# INPUTS: base_url (None → LLM_CONFIG.base_url), api_key (None → LLM_CONFIG.api_key).
# OUTPUTS: `OpenAI` / `AsyncOpenAI` с таймаутом OPENAI_TIMEOUT_SEC.
# KEYWORDS: factory, openai-compatible, timeout.
def make_llm_client(base_url: Optional[str] = None, api_key: Optional[str] = None) -> OpenAI:
    """
    Назначение:
        Создаёт синхронный клиент к OpenAI-совместимому сервису.

    Вход:
        base_url: URL сервиса до `/v1/`. `None` → `LLM_CONFIG.base_url` (env LLM_BASE_URL).
        api_key: Ключ. `None` → `LLM_CONFIG.api_key` (для локальных серверов — любое непустое значение).

    Выход:
        `OpenAI` с таймаутом `OPENAI_TIMEOUT_SEC`.
    """
    return OpenAI(
        base_url=base_url or LLM_CONFIG.base_url,
        api_key=api_key or LLM_CONFIG.api_key,
        timeout=OPENAI_TIMEOUT_SEC,
    )


def make_async_llm_client(base_url: Optional[str] = None, api_key: Optional[str] = None) -> AsyncOpenAI:
    """
    Назначение:
        Создаёт асинхронный клиент к OpenAI-совместимому сервису (используется VLM-парсером PDF).

    Вход:
        base_url, api_key: как у `make_llm_client`.

    Выход:
        `AsyncOpenAI` с таймаутом `OPENAI_TIMEOUT_SEC`.
    """
    return AsyncOpenAI(
        base_url=base_url or LLM_CONFIG.base_url,
        api_key=api_key or LLM_CONFIG.api_key,
        timeout=OPENAI_TIMEOUT_SEC,
    )
# END_CLIENT_FACTORY


# START_MODEL_RESOLVER
# PURPOSE: Резолвит имя LLM-модели. Нужно, потому что конфиги doc_types часто содержат стейл-идентификаторы облачных моделей (gpt-4.1-mini, openai/gpt-oss-120b), а локальный Spark-vLLM сервис ждёт имя, которое он зарегистрировал у себя.
# INPUTS: Имя модели из конфига (опционально).
# OUTPUTS: Имя модели, пригодное для Spark-vLLM.
# KEYWORDS: model-resolver, fallback, stale-cloud-id.
def resolve_runtime_llm_model(model_name: Optional[str]) -> str:
    """
    Назначение:
        Маппит стейл-идентификаторы облачных моделей (наследие от прежней конфигурации
        с облачными LLM) на имя модели, реально обслуживаемое Spark-vLLM. Дефолт —
        `LLM_CONFIG.default_model`.

    Вход:
        model_name: Имя модели (из конфига doc_type или вызывающего кода). Может быть None.

    Выход:
        Строка с именем модели для передачи в OpenAI-совместимый API.

    Логика:
        1. `None`/пусто → `LLM_CONFIG.default_model`.
        2. Имена с префиксом `openai/` или `gpt-` → `LLM_CONFIG.default_model`
           (эти имена были актуальны только для облачного API, Spark их не знает).
        3. Прочие имена передаются как есть.
    """
    fallback = LLM_CONFIG.default_model
    if not model_name:
        return fallback
    lowered = model_name.lower()
    if lowered.startswith("openai/") or lowered.startswith("gpt-"):
        return fallback
    return model_name
# END_MODEL_RESOLVER


# START_JSON_SANITIZER
# PURPOSE: Приватный хелпер: экранирует неэкранированные \\n/\\r/\\t внутри JSON-строк. LLM иногда возвращает сырые переводы строк внутри значений, что ломает `json.loads`.
def sanitize_json_string(s: str) -> str:
    """
    Назначение:
        Проходит по тексту посимвольно, отслеживает состояние «внутри JSON-строки»,
        заменяет сырые `\\n`, `\\r`, `\\t` на их экранированные варианты.

    Вход:
        s: Сырой текст ответа LLM.

    Выход:
        Строка с корректно экранированными управляющими символами внутри JSON-строк.

    Логика:
        1. `escape_next` — флаг, что предыдущий символ был `\\`.
        2. `in_string` — флаг, что мы сейчас внутри `"..."`.
        3. Встречая `"` в невэкранированном контексте — переключаем `in_string`.
        4. Встречая `\\n`/`\\r`/`\\t` ВНУТРИ строки — добавляем экранированную версию.
           Вне строки — оставляем как есть.
    """
    result_chars = []
    in_string = False
    escape_next = False
    for char in s:
        if escape_next:
            result_chars.append(char)
            escape_next = False
            continue
        if char == "\\" and in_string:
            result_chars.append(char)
            escape_next = True
            continue
        if char == '"':
            in_string = not in_string
            result_chars.append(char)
            continue
        if in_string and char in "\n\r\t":
            if char == "\n":
                result_chars.append("\\n")
            elif char == "\r":
                result_chars.append("\\r")
            elif char == "\t":
                result_chars.append("\\t")
        else:
            result_chars.append(char)
    return "".join(result_chars)
# END_JSON_SANITIZER


# START_JSON_PARSER
# PURPOSE: Парсинг JSON-ответа LLM в список нарушений. Понимает markdown-блоки, два формата ответа (`status=ok` и `status=fail + нарушения`).
# INPUTS: Сырой текст ответа, индекс и заголовок правила (для fallback-обогащения нарушений).
# OUTPUTS: Список нарушений. Пустой список при `status=ok` или ошибке парсинга.
# KEYWORDS: json-parse, markdown-strip, violations, status-ok.
def parse_json_response(raw_response: str, spec_index: int, spec_title: str) -> List[Dict[str, Any]]:
    """
    Назначение:
        Достаёт список нарушений из JSON-ответа LLM.

    Вход:
        raw_response: Сырой ответ LLM (может содержать markdown ```json … ``` обёртку).
        spec_index: Индекс правила — используется для обогащения нарушений, если LLM
            сам его не добавил.
        spec_title: Заголовок правила — то же самое для `rule_title`.

    Выход:
        Список dict'ов-нарушений. Если `status=ok` или JSON невалидный — пустой список.

    Логика:
        1. Убирает markdown-обёртку ```json ... ```, если есть.
        2. Чистит переносы строк внутри JSON-значений через `sanitize_json_string`.
        3. `json.loads`; при ошибке — предупреждение в stderr + пустой список.
        4. Если результат — dict со `status=ok` → пустой список.
        5. Если dict с `нарушения` → обогащает каждый `rule_index`/`rule_title`.
        6. Если список → возвращает как есть.
    """
    text = raw_response.strip()
    if text.startswith("```"):
        lines = text.split("\n")
        if lines[0].startswith("```"):
            lines = lines[1:]
        if lines and lines[-1].strip() == "```":
            lines = lines[:-1]
        text = "\n".join(lines)
    text = sanitize_json_string(text)
    try:
        result = json.loads(text)
        if isinstance(result, dict):
            if result.get("status") == "ok":
                return []
            violations = result.get("нарушения", [])
            for v in violations:
                v["rule_index"] = result.get("rule_index", spec_index)
                v["rule_title"] = result.get("rule_title", spec_title)
            return violations
        if isinstance(result, list):
            return result
        return []
    except json.JSONDecodeError as e:
        print(f"[WARN] Не удалось распарсить JSON: {e}", file=sys.stderr)
        print(f"[WARN] Ответ: {text[:200]}...", file=sys.stderr)
        return []
# END_JSON_PARSER


# START_CALL_LLM
# PURPOSE: Единственная функция, через которую весь рантайм вызывает LLM. Все параметры явные, `base_url=None` → облачный OpenAI из env.
# INPUTS: messages (OpenAI-совместимый формат), model, temperature, base_url, max_tokens, reasoning_effort, seed.
# OUTPUTS: Текст ответа LLM (`response.choices[0].message.content`).
# KEYWORDS: call-llm, openai, vllm, reasoning-effort, seed.
def call_llm(
    messages: List[Dict[str, str]],
    model: str,
    temperature: float = 0.0,
    base_url: Optional[str] = None,
    max_tokens: Optional[int] = None,
    reasoning_effort: Optional[str] = None,
    seed: Optional[int] = None,
    response_format: Optional[Dict[str, str]] = None,
) -> str:
    """
    Назначение:
        Выполняет один chat-completion запрос к LLM.

    Вход:
        messages: OpenAI-совместимый список `[{role, content}]`.
        model: Имя модели. Если содержит стейл-префикс (`openai/`, `gpt-`) — будет
            замаплен через `resolve_runtime_llm_model`.
        temperature: Температура генерации (0.0 = детерминизм).
        base_url: URL сервиса. `None` → облачный OpenAI через env-var `OPENAI_API_KEY`.
            Любое значение → Spark-vLLM с `api_key='none'`.
        max_tokens: Лимит токенов ответа. `None` → по умолчанию провайдера.
        reasoning_effort: Для gpt-oss-моделей: `low`/`medium`/`high`. Прокидывается
            через `extra_body={"reasoning_effort": ...}`.
        seed: Фиксированный seed для воспроизводимости (важно для MoE-моделей).
        response_format: Опциональный формат ответа (напр. `{"type": "json_object"}`).
            Нужен валидаторам KPSC/kartochka, которые требуют строгий JSON-вывод.

    Выход:
        Текст ответа LLM (или пустая строка, если `message.content` отсутствует).

    Логика:
        1. Если `base_url` задан — создаём OpenAI клиент для Spark-vLLM; иначе облачный.
           Клиент создаётся с `timeout=OPENAI_TIMEOUT_SEC` (защита от висящих запросов).
        2. Спускаем стейл-имена моделей через `resolve_runtime_llm_model` — только когда
           работаем с локальным сервисом (иначе облачный API знает свои имена).
        3. Собираем kwargs только с непустыми опциональными полями.
        4. ВСЕГДА добавляем в `extra_body` `chat_template_kwargs={'enable_thinking': False}`
           — иначе Qwen3.5 на Spark уходит в thinking-mode и обёртывает строгий JSON в
           markdown ` ```json ... ``` `, ломая `json.loads`. Мерджится с reasoning_effort.
        5. Делаем `chat.completions.create`, возвращаем content.
    """
    if base_url:
        model = resolve_runtime_llm_model(model)
        client = make_llm_client(base_url)
    else:
        # Облачный OpenAI: ключ берётся SDK из OPENAI_API_KEY. Единственный путь мимо фабрики.
        client = OpenAI(timeout=OPENAI_TIMEOUT_SEC)
    kwargs: Dict[str, Any] = {"model": model, "messages": messages, "temperature": temperature}
    if max_tokens is not None:
        kwargs["max_tokens"] = max_tokens
    if seed is not None:
        kwargs["seed"] = seed
    # extra_body всегда содержит chat_template_kwargs.enable_thinking=False;
    # reasoning_effort (если задан) добавляется поверх.
    extra_body: Dict[str, Any] = {"chat_template_kwargs": dict(DISABLE_THINKING_EXTRA_BODY["chat_template_kwargs"])}
    if reasoning_effort:
        extra_body["reasoning_effort"] = reasoning_effort
    kwargs["extra_body"] = extra_body
    if response_format is not None:
        kwargs["response_format"] = response_format
    response = client.chat.completions.create(**kwargs)
    return response.choices[0].message.content or ""
# END_CALL_LLM
