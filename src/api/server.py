# START_MODULE_CONTRACT
# PURPOSE: HTTP API-сервер аудита на FastAPI. Endpoints: /api/health, /api/login, /api/types, /api/audit (POST), /api/audit/<id>/events (SSE), /api/audit/<id>/result, /api/audit/<id>/download.
# INPUTS: HTTP-запросы (загрузка файла + doc_type). Обязательные переменные окружения: AUDIT_LOGIN, AUDIT_PASSWORD, AUDIT_TOKEN (без них сервер не стартует); необязательная CORS_ORIGINS (через запятую). Адреса LLM/OCR — из config/llm.py и config/parsers.py.
# OUTPUTS: app (FastAPI экземпляр) — уже подключённый со всеми middleware и роутами. Запускается через `uvicorn main:app` (main.py делает re-export).
# KEYWORDS: api, fastapi, sse, auth, queue, audit.
# LINKS: src/audit/engine.py (AuditEngine для multi_rule), src/engines.py (реестр спецдвижков и detect_engine).
# RATIONALE:
#   API-сервер — изолированный HTTP-слой. Реестр спецдвижков берётся из src/engines.py —
#   общего с CLI, чтобы набор типов в интерфейсе и в командной строке не расходился.
#   CUDA-warmup на импорт-тайме — намеренная инициализация GPU перед обработкой запросов.
# END_MODULE_CONTRACT

from __future__ import annotations

# START_IMPORTS
import asyncio
import faulthandler
import json
import os
import shutil
import sys
import threading
import traceback
import uuid
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional

from fastapi import FastAPI, File, Form, HTTPException, Request, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, JSONResponse
from pydantic import BaseModel
from sse_starlette.sse import EventSourceResponse

from config.llm import LLM_CONFIG
from config.parsers import PARSERS_CONFIG
from src.audit.engine import AuditEngine
from src.engines import SPECIAL_ENGINE_RUNNERS, detect_engine
from src.llm.client import make_llm_client
# END_IMPORTS


# START_PATHS
# PURPOSE: parents[2] = repo root (server.py лежит в src/api/).
_PROJECT_ROOT = Path(__file__).resolve().parents[2]
_DOC_CONFIGS_DIR = _PROJECT_ROOT / "doc_configs"
_LOGS_RESULT_DIR = _PROJECT_ROOT / "logs_result"
UPLOAD_DIR = _PROJECT_ROOT / "uploads"
UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
# END_PATHS


# START_CUDA_WARMUP
# PURPOSE: Прогрев CUDA-контекста при импорте, чтобы первый запрос не тратил время на
# инициализацию GPU. Если torch не установлен (модели на удалённом сервере) —
# пропускаем без ошибки.
faulthandler.enable(file=sys.stderr)

try:
    import torch
    if torch.cuda.is_available():
        torch.cuda.init()
        _dummy = torch.zeros(1, device="cuda:0")
        del _dummy
        print(f"[CUDA] Контекст зарезервирован на {torch.cuda.get_device_name(0)}")
    else:
        print("[CUDA] GPU не обнаружен — пропуск CUDA warmup (модели на удалённом сервере)")
except ImportError:
    print("[CUDA] torch не установлен — пропуск CUDA warmup (модели на удалённом сервере)")
except Exception as e:
    print(f"[CUDA] Не удалось зарезервировать контекст: {e}")
# END_CUDA_WARMUP


# START_APP
app = FastAPI(title="AuditAPI")

# Разрешённые origin фронтенда — из CORS_ORIGINS (через запятую). По умолчанию — локальный preview :5174.
_CORS_ORIGINS = [o.strip() for o in os.getenv("CORS_ORIGINS", "http://localhost:5174,http://127.0.0.1:5174").split(",") if o.strip()]

app.add_middleware(
    CORSMiddleware,
    allow_origins=_CORS_ORIGINS,
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)
# END_APP


# START_AUTH
# PURPOSE: Cookie/Bearer-token аутентификация. Логин, пароль и токен — ТОЛЬКО из окружения
# (AUDIT_LOGIN / AUDIT_PASSWORD / AUDIT_TOKEN, см. .env.example). Дефолтов в коде нет намеренно:
# без них сервер при старте (uvicorn) падает с понятным сообщением. Импорт модуля при этом
# не ломается — CLI и тесты работают без этих переменных.
AUTH_LOGIN = os.getenv("AUDIT_LOGIN", "")
AUTH_PASSWORD = os.getenv("AUDIT_PASSWORD", "")
AUTH_TOKEN = os.getenv("AUDIT_TOKEN", "")
_REQUIRED_AUTH_ENV = ("AUDIT_LOGIN", "AUDIT_PASSWORD", "AUDIT_TOKEN")


def check_auth_env() -> None:
    """Проверяет, что все переменные аутентификации заданы. Бросает RuntimeError с перечнем пустых."""
    missing = [name for name in _REQUIRED_AUTH_ENV if not os.getenv(name)]
    if missing:
        raise RuntimeError(
            "Не заданы переменные окружения: " + ", ".join(missing)
            + ". Задайте их в .env (образец — .env.example) или в окружении процесса."
        )


class LoginRequest(BaseModel):
    """Тело запроса авторизации."""
    login: str
    password: str


def _check_auth(request: Request) -> None:
    """Проверяет токен в cookie или в Authorization-header. Бросает 401 если невалиден."""
    token = request.cookies.get("auth_token")
    if not token:
        auth_header = request.headers.get("Authorization", "")
        if auth_header.startswith("Bearer "):
            token = auth_header[7:]
    if not AUTH_TOKEN or token != AUTH_TOKEN:
        raise HTTPException(status_code=401, detail="Не авторизован")
# END_AUTH


# START_QUEUE_AND_SESSIONS
# PURPOSE: Глобальная очередь аудитов (одно одновременно через `_audit_lock`)
# + словарь активных сессий.
_audit_lock = threading.Lock()
_queue_counter = 0
_queue_counter_lock = threading.Lock()
sessions: Dict[str, Dict[str, Any]] = {}
# END_QUEUE_AND_SESSIONS


# START_DISPATCH
# PURPOSE: Реестр спецдвижков и detect_engine — из src/engines.py (общие с CLI).
# END_DISPATCH


_STANDARD_EXTENSIONS = {".docx", ".doc", ".pdf", ".odt", ".rtf"}
_XLSX_EXTENSIONS = {".xlsx"}
_PPTX_EXTENSIONS = {".pptx"}


def _get_allowed_extensions(doc_type: str) -> set:
    """
    Назначение:
        Возвращает допустимые расширения файлов для данного doc_type.

    Логика:
        - Спецдвижки (движок из SPECIAL_ENGINE_RUNNERS) → .xlsx (все они читают Excel)
        - Остальные → берём из parser_by_ext конфига: тип принимает ровно то,
          что реально умеет распарсить (иначе файл упал бы уже в середине аудита)
        - parser: "pptx" (старый формат конфига) → .pptx
        - secondary_file → плюс .{secondary.type}
        - Конфига/parser_by_ext нет → стандартные (.docx/.doc/.pdf/.odt/.rtf)
    """
    config_path = _DOC_CONFIGS_DIR / doc_type / "config.json"
    if not config_path.exists():
        return _STANDARD_EXTENSIONS
    with open(config_path, "r", encoding="utf-8") as f:
        data = json.load(f)
    engine = data.get("engine", "standard")
    if engine in SPECIAL_ENGINE_RUNNERS:
        return _XLSX_EXTENSIONS
    if data.get("parser") == "pptx":
        return _PPTX_EXTENSIONS
    allowed = {e.lower() for e in (data.get("parser_by_ext") or {})}
    if not allowed:
        allowed = set(_STANDARD_EXTENSIONS)
    if data.get("secondary_file"):
        sec_type = data["secondary_file"].get("type", "")
        if sec_type:
            allowed.add(f".{sec_type}")
    return allowed


def _run_special_engine(engine_type: str, doc_type: str, target_path: str, session_dir: str):
    """
    Запуск спецдвижка через прямую callable из `SPECIAL_ENGINE_RUNNERS`.
    Спецдвижки ожидают argparse-Namespace — эмулируем его.
    """
    from types import SimpleNamespace
    runner = SPECIAL_ENGINE_RUNNERS[engine_type]
    args = SimpleNamespace(
        target=target_path, doc_type=doc_type, parse_only=False, rule_filter=None,
        model=None, temperature=None, session_dir=session_dir,
    )
    return runner(args)


def _run_audit_thread(session_id: str, doc_type: str, target_path: str) -> None:
    """
    Запуск аудита в фоновом потоке с очередью и progress_callback.
    Обновляет `sessions[session_id]` по мере выполнения.
    """
    global _queue_counter
    session = sessions[session_id]
    loop = session["loop"]
    queue = session["events"]

    def progress_callback(event_type: str, data: dict) -> None:
        """Колбэк из engine — пушит события в SSE-очередь."""
        asyncio.run_coroutine_threadsafe(queue.put({"type": event_type, "data": data}), loop)

    with _queue_counter_lock:
        _queue_counter += 1
        position = _queue_counter
    if position > 1:
        progress_callback("queue", {"position": position - 1})
        print(f"[QUEUE] Сессия {session_id} встала в очередь, позиция {position - 1}")
    try:
        with _audit_lock:
            progress_callback("queue", {"position": 0})
            try:
                engine_type = detect_engine(doc_type)
                if engine_type in SPECIAL_ENGINE_RUNNERS:
                    progress_callback("audit_start", {
                        "doc_type": doc_type, "filename": Path(target_path).name, "engine": engine_type,
                    })
                    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                    session_dir = str(_LOGS_RESULT_DIR / doc_type / f"session_{timestamp}")
                    result = _run_special_engine(engine_type, doc_type, target_path, session_dir)
                else:
                    engine = AuditEngine(doc_type)
                    result = engine.run(target_path, progress_callback=progress_callback)
                session["result"] = {
                    "violations": result.violations,
                    "rules_checked": result.rules_checked,
                    "duration_sec": result.duration_sec,
                    "warnings": result.warnings,
                    "unchecked": result.unchecked_rules,
                }
                session["session_dir"] = str(result.session_dir)
                session["status"] = "done"
                progress_callback("complete", session["result"])
            except Exception as e:
                err_msg = f"{type(e).__name__}: {e}"
                traceback.print_exc(file=sys.stderr)
                sys.stderr.flush()
                print(f"[AUDIT ERROR] {err_msg}", file=sys.stderr, flush=True)
                session["status"] = "error"
                session["error_message"] = err_msg
                # Событие называется audit_error (НЕ error): у EventSource на фронте имя error
                # совпадает с DOM-событием обрыва соединения — см. api.js::subscribeToProgress.
                progress_callback("audit_error", {"message": err_msg})
    finally:
        try:
            sd = session.get("session_dir")
            if sd:
                orig_dir = Path(sd) / "original"
                orig_dir.mkdir(exist_ok=True)
                shutil.copy2(target_path, orig_dir / Path(target_path).name)
        except Exception:
            pass
        with _queue_counter_lock:
            _queue_counter -= 1
# END_DISPATCH


# START_ENDPOINTS
@app.on_event("startup")
async def _require_auth_env() -> None:
    """При старте сервера требуем AUDIT_* в окружении — иначе падаем с понятным сообщением."""
    check_auth_env()


@app.get("/api/health")
async def health_check():
    """Проверка здоровья всех сервисов. Без авторизации. Адреса — из конфига (env LLM_BASE_URL / OCR_BASE_URL)."""
    from urllib.request import urlopen
    services = {}
    ocr_models_url = PARSERS_CONFIG.pdf.vlm.base_url.rstrip("/") + "/models"
    llm_models_url = LLM_CONFIG.base_url.rstrip("/") + "/models"
    try:
        r = urlopen(ocr_models_url, timeout=5)
        services["paddleocr"] = {"status": "ok", "code": r.status, "url": ocr_models_url}
    except Exception as e:
        services["paddleocr"] = {"status": "error", "detail": str(e), "url": ocr_models_url}
    try:
        r = urlopen(llm_models_url, timeout=5)
        services["llm"] = {"status": "ok", "code": r.status, "url": llm_models_url}
    except Exception as e:
        services["llm"] = {"status": "error", "detail": str(e), "url": llm_models_url}
    try:
        import torch as _torch
        cuda_ok = _torch.cuda.is_available()
        services["cuda"] = {
            "status": "ok" if cuda_ok else "skip",
            "device": _torch.cuda.get_device_name(0) if cuda_ok else "нет GPU (модели на удалённом сервере)",
        }
    except ImportError:
        services["cuda"] = {"status": "skip", "device": "torch не установлен (модели на удалённом сервере)"}
    except Exception as e:
        services["cuda"] = {"status": "error", "detail": str(e)}
    overall = all((s["status"] in ("ok", "skip") for s in services.values()))
    return {
        "status": "ok" if overall else "degraded",
        "services": services,
        "timestamp": datetime.now().isoformat(),
    }


@app.post("/api/login")
async def login(body: LoginRequest):
    """Авторизация. Возвращает token и ставит cookie."""
    valid_logins = {AUTH_LOGIN, "guest"}
    if body.login not in valid_logins or body.password != AUTH_PASSWORD:
        raise HTTPException(status_code=401, detail="Неверный логин или пароль")
    response = JSONResponse({"ok": True, "token": AUTH_TOKEN})
    response.set_cookie("auth_token", AUTH_TOKEN, httponly=True, max_age=2592000, samesite="lax")
    return response


import re as _re


def _title_code_key(t):
    """Ключ сортировки по коду подтипа, уже присутствующему в начале doc_title (напр. «3.10 …»)."""
    m = _re.match(r"\s*(\d+)\.(\d+)", t.get("doc_title", ""))
    return (int(m.group(1)), int(m.group(2))) if m else (99, 99)


@app.get("/api/types")
async def get_types(request: Request):
    """Список типов документов, отсортированный по коду подтипа (0.1 → 3.10)."""
    _check_auth(request)
    types = AuditEngine.list_doc_types()
    types.sort(key=_title_code_key)
    return types


# ── Пользовательские правила (overlay rules_custom.json) — только generic-LLM типы ──
CUSTOM_INDEX_START = 200

# START_AUTHOR_PROMPT
_RULE_AUTHOR_SYSTEM_PROMPT = (
    "Ты — методист-редактор правил проверки документов. По черновой формулировке эксперта "
    "составь ЧЁТКУЮ, ОБЪЕКТИВНУЮ инструкцию проверки указанной части документа для автоматической сверки.\n"
    "Требования к инструкции (поле check):\n"
    "- императив, лаконично, на русском;\n"
    "- перечисли конкретные проверяемые критерии; если критериев несколько — добавь «(каждое несоблюдение — нарушение)»;\n"
    "- проверяй НАЛИЧИЕ и ФОРМАТ реквизитов, НЕ выдумывай конкретных значений (ФИО, номера, даты — любые);\n"
    "- не ссылайся на другие секции.\n"
    "Поле exclusions — что НЕ считать нарушением (может быть пустой строкой).\n"
    "Верни СТРОГО JSON без пояснений и markdown: {\"check\": \"...\", \"exclusions\": \"...\"}"
)
# END_AUTHOR_PROMPT


def _type_engine(doc_type: str) -> str:
    cfg = _DOC_CONFIGS_DIR / doc_type / "config.json"
    if not cfg.exists():
        return "vision"
    try:
        return json.loads(cfg.read_text(encoding="utf-8")).get("engine", "vision")
    except (json.JSONDecodeError, OSError):
        return "vision"


def _type_is_editable(doc_type: str) -> bool:
    """Правила можно добавлять только у generic-LLM типа (engine=vision) с базовыми правилами."""
    return _type_engine(doc_type) == "vision" and (_DOC_CONFIGS_DIR / doc_type / "rules_multi.json").exists()


def _load_sections_map(doc_type: str) -> Dict[str, Any]:
    p = _DOC_CONFIGS_DIR / doc_type / "sections.json"
    if not p.exists():
        return {}
    try:
        return json.loads(p.read_text(encoding="utf-8")).get("sections", {})
    except (json.JSONDecodeError, OSError):
        return {}


def _load_base_rules(doc_type: str) -> list:
    p = _DOC_CONFIGS_DIR / doc_type / "rules_multi.json"
    if not p.exists():
        return []
    data = json.loads(p.read_text(encoding="utf-8"))
    return data["rules"] if isinstance(data, dict) else data


def _custom_rules_path(doc_type: str) -> Path:
    return _DOC_CONFIGS_DIR / doc_type / "rules_custom.json"


def _load_custom_rules(doc_type: str) -> list:
    p = _custom_rules_path(doc_type)
    if not p.exists():
        return []
    try:
        return json.loads(p.read_text(encoding="utf-8")).get("rules", [])
    except (json.JSONDecodeError, OSError):
        return []


def _save_custom_rules(doc_type: str, rules: list) -> None:
    """Атомарная запись overlay: temp-файл + replace."""
    p = _custom_rules_path(doc_type)
    tmp = p.with_name(p.name + ".tmp")
    tmp.write_text(json.dumps({"rules": rules}, ensure_ascii=False, indent=2), encoding="utf-8")
    tmp.replace(p)


def _serialize_rule(r: dict, sections_map: dict, is_custom: bool) -> dict:
    # нормализация: target_section (строка) -> [target_section]; target_sections — как есть
    if "target_sections" in r:
        targets = r["target_sections"]
    elif "target_section" in r:
        targets = [r["target_section"]]
    else:
        targets = []
    sections = [
        {"name": key, "description": sections_map.get(key, {}).get("description", "")}
        for key in targets
    ]
    return {
        "index": r.get("index"),
        "title": r.get("title"),
        "check": r.get("check"),
        "exclusions": r.get("exclusions", ""),
        "target_sections": targets,
        "sections": sections,
        "layer": "base",
        "custom": is_custom,
    }


def _extract_json(text: str) -> dict:
    """Достаёт JSON-объект из ответа модели (снимает ```-ограждения, берёт первый {...})."""
    t = text.strip()
    if t.startswith("```"):
        t = t.strip("`")
        if t.lstrip().lower().startswith("json"):
            t = t.lstrip()[4:]
    start = t.find("{")
    end = t.rfind("}")
    if start != -1 and end != -1 and end > start:
        t = t[start:end + 1]
    return json.loads(t)


def _author_rule(doc_type: str, title: str, raw_check: str, sections_sel: List[str], sections_map: dict) -> dict:
    """Спец-модель переписывает сырьё эксперта в чёткую инструкцию (check + exclusions)."""
    cfg = json.loads((_DOC_CONFIGS_DIR / doc_type / "config.json").read_text(encoding="utf-8"))
    base_url = cfg.get("llm_base_url") or LLM_CONFIG.base_url
    model = cfg.get("model") or LLM_CONFIG.default_model
    doc_title = cfg.get("doc_title", doc_type)
    # Контекст выбранных секций (одна, несколько или все = «весь документ»).
    all_sections = [s for s in sections_map if s != "filename"]
    whole_doc = len(sections_sel) >= len(all_sections) and len(sections_sel) > 1
    sec_lines = "; ".join(f"«{s}» — {sections_map.get(s, {}).get('description', '')}" for s in sections_sel)
    where = "ВЕСЬ ДОКУМЕНТ (все секции): " + sec_lines if whole_doc else "Секции проверки: " + sec_lines
    user = (
        f"Тип документа: {doc_title}\n"
        f"{where}\n"
        f"Заголовок правила: {title}\n"
        f"Черновая формулировка эксперта (что проверять): {raw_check}\n\n"
        "Составь строгую инструкцию проверки для указанных секций и верни JSON {check, exclusions}. "
        "Если проверка охватывает несколько секций или весь документ — учти это в формулировке."
    )
    client = make_llm_client(base_url)
    resp = client.chat.completions.create(
        model=model,
        messages=[
            {"role": "system", "content": _RULE_AUTHOR_SYSTEM_PROMPT},
            {"role": "user", "content": user},
        ],
        temperature=0.0,
        max_tokens=2048,
        extra_body={"chat_template_kwargs": {"enable_thinking": False}},
    )
    text = resp.choices[0].message.content or ""
    try:
        data = _extract_json(text)
    except (json.JSONDecodeError, ValueError):
        data = {}
    return {
        "title": title,
        "check": (data.get("check") or raw_check).strip(),
        "exclusions": (data.get("exclusions") or "").strip(),
        "target_sections": sections_sel,
    }


class RuleDraftRequest(BaseModel):
    """Сырьё эксперта для формулировки правила спец-моделью."""
    title: str
    raw_check: str
    sections: List[str]


class RuleSaveRequest(BaseModel):
    """Готовое (подтверждённое) правило для сохранения/изменения."""
    title: str
    check: str
    target_sections: List[str]
    exclusions: Optional[str] = ""


@app.get("/api/types/{doc_type}/rules")
async def get_type_rules(doc_type: str, request: Request):
    """Правила типа: база (view-only) + кастом (overlay). editable=true только для generic-LLM типа."""
    _check_auth(request)
    if not (_DOC_CONFIGS_DIR / doc_type / "rules_multi.json").exists():
        # special-runner или тип без текстовых правил — редактирование недоступно
        return {"editable": False, "sections": [], "rules": []}
    sections_map = _load_sections_map(doc_type)
    sections_list = [
        {"name": k, "description": v.get("description", ""), "start": v.get("start", ""), "end": v.get("end", "")}
        for k, v in sections_map.items() if k != "filename"
    ]
    base = [_serialize_rule(r, sections_map, False) for r in _load_base_rules(doc_type)]
    custom = [_serialize_rule(r, sections_map, True) for r in _load_custom_rules(doc_type)]
    return {"editable": _type_is_editable(doc_type), "sections": sections_list, "rules": base + custom}


@app.post("/api/types/{doc_type}/rules/draft")
async def draft_type_rule(doc_type: str, body: RuleDraftRequest, request: Request):
    """Спец-модель формулирует правило из сырья эксперта. Ничего не пишет на диск."""
    _check_auth(request)
    if not _type_is_editable(doc_type):
        raise HTTPException(status_code=400, detail="Для этого типа нельзя добавлять правила (проверка задана алгоритмом)")
    sections_map = _load_sections_map(doc_type)
    if not body.sections or any(s not in sections_map for s in body.sections):
        raise HTTPException(status_code=400, detail="Неизвестная секция")
    # Вызов модели синхронный и долгий — выносим в поток, чтобы не блокировать event-loop
    # (иначе одна зависшая «Сформулировать» замораживает health/SSE всем, замер 2.10c).
    return await asyncio.to_thread(
        _author_rule, doc_type, body.title.strip(), body.raw_check.strip(), body.sections, sections_map
    )


@app.post("/api/types/{doc_type}/rules")
async def create_type_rule(doc_type: str, body: RuleSaveRequest, request: Request):
    """Сохранить пользовательское правило в overlay rules_custom.json (индекс с 200)."""
    _check_auth(request)
    if not _type_is_editable(doc_type):
        raise HTTPException(status_code=400, detail="Для этого типа нельзя добавлять правила")
    sections_map = _load_sections_map(doc_type)
    if not body.target_sections or any(s not in sections_map for s in body.target_sections):
        raise HTTPException(status_code=400, detail="Неизвестная секция")
    custom = _load_custom_rules(doc_type)
    next_index = max([r.get("index", 0) for r in custom] + [CUSTOM_INDEX_START - 1]) + 1
    rule = {
        "index": next_index,
        "title": body.title.strip(),
        "target_sections": body.target_sections,
        "check": body.check.strip(),
        "custom": True,
    }
    if body.exclusions:
        rule["exclusions"] = body.exclusions.strip()
    custom.append(rule)
    _save_custom_rules(doc_type, custom)
    return _serialize_rule(rule, sections_map, True)


@app.put("/api/types/{doc_type}/rules/{index}")
async def update_type_rule(doc_type: str, index: int, body: RuleSaveRequest, request: Request):
    """Изменить пользовательское правило (только index ≥ 200; база не редактируется)."""
    _check_auth(request)
    if index < CUSTOM_INDEX_START:
        raise HTTPException(status_code=400, detail="Редактировать можно только пользовательские правила")
    sections_map = _load_sections_map(doc_type)
    if not body.target_sections or any(s not in sections_map for s in body.target_sections):
        raise HTTPException(status_code=400, detail="Неизвестная секция")
    custom = _load_custom_rules(doc_type)
    for r in custom:
        if r.get("index") == index:
            r["title"] = body.title.strip()
            r["target_sections"] = body.target_sections
            r["check"] = body.check.strip()
            r.pop("target_section", None)
            if body.exclusions:
                r["exclusions"] = body.exclusions.strip()
            else:
                r.pop("exclusions", None)
            _save_custom_rules(doc_type, custom)
            return _serialize_rule(r, sections_map, True)
    raise HTTPException(status_code=404, detail="Правило не найдено")


@app.delete("/api/types/{doc_type}/rules/{index}")
async def delete_type_rule(doc_type: str, index: int, request: Request):
    """Удалить пользовательское правило (только index ≥ 200)."""
    _check_auth(request)
    if index < CUSTOM_INDEX_START:
        raise HTTPException(status_code=400, detail="Удалять можно только пользовательские правила")
    custom = _load_custom_rules(doc_type)
    new_rules = [r for r in custom if r.get("index") != index]
    if len(new_rules) == len(custom):
        raise HTTPException(status_code=404, detail="Правило не найдено")
    _save_custom_rules(doc_type, new_rules)
    return {"ok": True}


@app.post("/api/audit")
async def start_audit(request: Request, file: UploadFile = File(...), doc_type: str = Form(...)):
    """Запуск аудита: принимает файл + тип, возвращает session_id."""
    _check_auth(request)
    file_ext = Path(file.filename).suffix.lower() if file.filename else ""
    allowed = _get_allowed_extensions(doc_type)
    if file_ext not in allowed:
        ext_list = ", ".join(sorted(allowed))
        raise HTTPException(
            status_code=400,
            detail=f"Неподдерживаемый формат файла «{file_ext}». Для типа «{doc_type}» допустимы: {ext_list}",
        )
    session_id = uuid.uuid4().hex[:8]
    upload_path = UPLOAD_DIR / session_id / file.filename
    upload_path.parent.mkdir(parents=True, exist_ok=True)
    with open(upload_path, "wb") as f:
        shutil.copyfileobj(file.file, f)
    event_queue = asyncio.Queue()
    loop = asyncio.get_event_loop()
    sessions[session_id] = {
        "status": "running", "events": event_queue, "loop": loop,
        "result": None, "doc_type": doc_type, "filename": file.filename,
        "upload_path": str(upload_path), "session_dir": None,
    }
    thread = threading.Thread(
        target=_run_audit_thread, args=(session_id, doc_type, str(upload_path)), daemon=True,
    )
    thread.start()
    return {"session_id": session_id}


@app.post("/api/audit/cross")
async def start_cross_audit(request: Request, files: List[UploadFile] = File(...)):
    """Запуск сквозной сверки 3 документов (2.4 Карточка + 0.6 Протокол + 0.5 Приказ о тираже)."""
    _check_auth(request)
    doc_type = "crosscheck_2_4_0_6_0_5"
    if not files or len(files) < 2:
        raise HTTPException(
            status_code=400,
            detail="Нужно загрузить 3 документа: 2.4 Карточка проекта, 0.6 Протокол выполнения, 0.5 Приказ о тираже",
        )
    session_id = uuid.uuid4().hex[:8]
    folder = UPLOAD_DIR / session_id
    folder.mkdir(parents=True, exist_ok=True)
    names = []
    for f in files:
        dest = folder / (f.filename or uuid.uuid4().hex)
        with open(dest, "wb") as out:
            shutil.copyfileobj(f.file, out)
        names.append(f.filename)
    event_queue = asyncio.Queue()
    loop = asyncio.get_event_loop()
    sessions[session_id] = {
        "status": "running", "events": event_queue, "loop": loop,
        "result": None, "doc_type": doc_type, "filename": ", ".join(names),
        "upload_path": str(folder), "session_dir": None,
    }
    thread = threading.Thread(
        target=_run_audit_thread, args=(session_id, doc_type, str(folder)), daemon=True,
    )
    thread.start()
    return {"session_id": session_id}


@app.get("/api/audit/{session_id}/events")
async def audit_events(session_id: str):
    """SSE-стрим прогресса аудита."""
    session = sessions.get(session_id)
    if not session:
        raise HTTPException(status_code=404, detail="Сессия не найдена")

    async def event_generator():
        queue = session["events"]
        while True:
            try:
                event = await asyncio.wait_for(queue.get(), timeout=300)
            except asyncio.TimeoutError:
                yield {"event": "ping", "data": "{}"}
                continue
            yield {"event": event["type"], "data": json.dumps(event["data"], ensure_ascii=False)}
            if event["type"] in ("complete", "audit_error"):
                break

    return EventSourceResponse(event_generator())


@app.get("/api/audit/{session_id}/result")
async def get_result(request: Request, session_id: str):
    """Получить результат аудита (violations + статистика)."""
    _check_auth(request)
    session = sessions.get(session_id)
    if not session:
        raise HTTPException(status_code=404, detail="Сессия не найдена")
    if session["status"] == "running":
        raise HTTPException(status_code=202, detail="Аудит ещё выполняется")
    if session["status"] == "error":
        raise HTTPException(status_code=500, detail=f"Аудит завершился с ошибкой: {session.get('error_message', '')}".rstrip(": "))
    return session["result"]


@app.get("/api/audit/{session_id}/download")
async def download_report(request: Request, session_id: str):
    """Скачать Excel-отчёт."""
    _check_auth(request)
    session = sessions.get(session_id)
    if not session or not session.get("session_dir"):
        raise HTTPException(status_code=404, detail="Отчёт не найден")
    session_dir = Path(session["session_dir"])
    xlsx_path = session_dir / "audit_result.xlsx"
    if not xlsx_path.exists():
        xlsx_files = list(session_dir.glob("*.xlsx"))
        if not xlsx_files:
            raise HTTPException(status_code=404, detail="Excel не сформирован")
        xlsx_path = xlsx_files[0]
    return FileResponse(
        str(xlsx_path),
        filename=f"audit_{session_id}.xlsx",
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
# END_ENDPOINTS
