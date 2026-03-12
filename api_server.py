#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
FastAPI-сервер для веб-аудита документов.

Эндпоинты:
  POST /api/login             — авторизация (логин/пароль из env)
  GET  /api/types             — список типов документов
  POST /api/audit             — запуск аудита (multipart: file + doc_type)
  GET  /api/audit/{id}/events — SSE-стрим прогресса
  GET  /api/audit/{id}/download — скачать Excel-отчёт

Запуск:
  uvicorn api_server:app --host 0.0.0.0 --port 8080 --reload
"""

import asyncio
import faulthandler
import json
import os
import shutil
import sys
import traceback

# Печатает traceback даже при segfault (SIGSEGV, SIGFPE, SIGABRT)
faulthandler.enable(file=sys.stderr)

# Резервируем CUDA-контекст при старте (~1.5GB), ДО запуска VLLM.
# Без этого VLLM забирает всю GPU-память и LayoutDetector не может инициализироваться.
try:
    import torch
    if torch.cuda.is_available():
        torch.cuda.init()
        _dummy = torch.zeros(1, device="cuda:0")
        del _dummy
        print(f"[CUDA] Контекст зарезервирован на {torch.cuda.get_device_name(0)}")
except Exception as e:
    print(f"[CUDA] Не удалось зарезервировать контекст: {e}")
import threading
import uuid
from pathlib import Path
from typing import Any, Dict, Optional

from fastapi import FastAPI, File, Form, HTTPException, Request, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
from pydantic import BaseModel
from sse_starlette.sse import EventSourceResponse

app = FastAPI(title="AuditAPI")

# Глобальная блокировка: один аудит за раз, остальные в очереди
_audit_lock = threading.Lock()
_queue_counter = 0  # Сколько потоков ждут/выполняют аудит
_queue_counter_lock = threading.Lock()  # Защита счётчика

# CORS для dev-режима (Vite на :5173)
app.add_middleware(
    CORSMiddleware,
    allow_origins=["http://localhost:5173", "http://192.168.20.118:5173"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# --- Конфигурация ---
AUTH_LOGIN = os.getenv("AUDIT_LOGIN", "admin")
AUTH_PASSWORD = os.getenv("AUDIT_PASSWORD", "mos186124kva")
AUTH_TOKEN = os.getenv("AUDIT_TOKEN", "audit-session-token-2026")
UPLOAD_DIR = Path("uploads")
UPLOAD_DIR.mkdir(exist_ok=True)

# --- In-memory хранилище сессий ---
sessions: Dict[str, Dict[str, Any]] = {}


class LoginRequest(BaseModel):
    """Тело запроса авторизации."""
    login: str
    password: str


def _check_auth(request: Request):
    """Проверка авторизации через cookie или header."""
    token = request.cookies.get("auth_token")
    if not token:
        auth_header = request.headers.get("Authorization", "")
        if auth_header.startswith("Bearer "):
            token = auth_header[7:]
    if token != AUTH_TOKEN:
        raise HTTPException(status_code=401, detail="Не авторизован")


# ── POST /api/login ──────────────────────────────────────────────

@app.post("/api/login")
async def login(body: LoginRequest):
    """Авторизация. Возвращает token и ставит cookie."""
    if body.login != AUTH_LOGIN or body.password != AUTH_PASSWORD:
        raise HTTPException(status_code=401, detail="Неверный логин или пароль")

    from fastapi.responses import JSONResponse
    response = JSONResponse({"ok": True, "token": AUTH_TOKEN})
    response.set_cookie(
        "auth_token", AUTH_TOKEN,
        httponly=True, max_age=86400, samesite="lax"
    )
    return response


# ── GET /api/types ────────────────────────────────────────────────

@app.get("/api/types")
async def get_types(request: Request):
    """Список доступных типов документов."""
    _check_auth(request)
    from audit_engine.engine import AuditEngine
    return AuditEngine.list_doc_types()


# ── POST /api/audit ───────────────────────────────────────────────

@app.post("/api/audit")
async def start_audit(
    request: Request,
    file: UploadFile = File(...),
    doc_type: str = Form(...)
):
    """Запуск аудита: принимает файл + тип, возвращает session_id."""
    _check_auth(request)

    session_id = uuid.uuid4().hex[:8]

    # Сохраняем файл
    upload_path = UPLOAD_DIR / session_id / file.filename
    upload_path.parent.mkdir(parents=True, exist_ok=True)
    with open(upload_path, "wb") as f:
        shutil.copyfileobj(file.file, f)

    # Создаём очередь SSE-событий
    event_queue = asyncio.Queue()
    loop = asyncio.get_event_loop()

    sessions[session_id] = {
        "status": "running",
        "events": event_queue,
        "loop": loop,
        "result": None,
        "doc_type": doc_type,
        "filename": file.filename,
        "upload_path": str(upload_path),
        "session_dir": None,
    }

    # Запуск аудита в фоновом потоке
    thread = threading.Thread(
        target=_run_audit_thread,
        args=(session_id, doc_type, str(upload_path)),
        daemon=True,
    )
    thread.start()

    return {"session_id": session_id}


# ── GET /api/audit/{id}/events (SSE) ─────────────────────────────

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
                # Keepalive
                yield {"event": "ping", "data": "{}"}
                continue

            yield {
                "event": event["type"],
                "data": json.dumps(event["data"], ensure_ascii=False),
            }
            if event["type"] in ("complete", "error"):
                break

    return EventSourceResponse(event_generator())


# ── GET /api/audit/{id}/result ────────────────────────────────────

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
        raise HTTPException(status_code=500, detail="Аудит завершился с ошибкой")
    return session["result"]


# ── GET /api/audit/{id}/download ──────────────────────────────────

@app.get("/api/audit/{session_id}/download")
async def download_report(request: Request, session_id: str):
    """Скачать Excel-отчёт."""
    _check_auth(request)
    session = sessions.get(session_id)
    if not session or not session.get("session_dir"):
        raise HTTPException(status_code=404, detail="Отчёт не найден")

    session_dir = Path(session["session_dir"])
    # Стандартный pipeline → audit_result.xlsx
    # Спецдвижки → validation_report.xlsx (kpsc, kartochka) или 03_report.xlsx (drivers)
    xlsx_path = session_dir / "audit_result.xlsx"
    if not xlsx_path.exists():
        # Ищем любой xlsx в директории сессии
        xlsx_files = list(session_dir.glob("*.xlsx"))
        if not xlsx_files:
            raise HTTPException(status_code=404, detail="Excel не сформирован")
        xlsx_path = xlsx_files[0]

    return FileResponse(
        str(xlsx_path),
        filename=f"audit_{session_id}.xlsx",
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


# ── Фоновый поток аудита ─────────────────────────────────────────

def _detect_engine(doc_type: str) -> str:
    """
    Определяет тип движка по полю "engine" в config.json.

    Returns:
        str: значение поля "engine" или "standard" если поля нет.
    """
    config_path = Path(__file__).parent / "doc_configs" / doc_type / "config.json"
    if config_path.exists():
        with open(config_path, "r", encoding="utf-8") as f:
            data = json.load(f)
        return data.get("engine", "standard")
    return "standard"


# Маппинг спецдвижков на модули (аналог run_audit.py)
SPECIAL_ENGINES = {
    "drivers": "audit_engine.drivers",
    "kpsc": "audit_engine.kpsc",
    "kartochka_proekta": "audit_engine.kartochka_proekta",
}


def _run_special_engine(engine_type: str, doc_type: str, target_path: str,
                        session_dir: str) -> "AuditResult":
    """
    Запуск спецдвижка (kpsc, drivers, kartochka_proekta).

    Спецдвижки ожидают argparse Namespace — эмулируем его.

    Args:
        engine_type: ключ из SPECIAL_ENGINES
        doc_type: тип документа
        target_path: путь к файлу
        session_dir: директория для логов
    Returns:
        AuditResult
    """
    import importlib
    from types import SimpleNamespace

    module = importlib.import_module(SPECIAL_ENGINES[engine_type])
    args = SimpleNamespace(
        target=target_path,
        parse_only=False,
        rule_filter=None,
        model=None,
        temperature=None,
        session_dir=session_dir,
    )
    return module.run(args)


def _run_audit_thread(session_id: str, doc_type: str, target_path: str):
    """Запуск аудита в фоновом потоке с очередью и progress_callback."""
    global _queue_counter
    session = sessions[session_id]
    loop = session["loop"]
    queue = session["events"]

    def progress_callback(event_type: str, data: dict):
        """Колбэк из engine — пушит события в SSE-очередь."""
        asyncio.run_coroutine_threadsafe(
            queue.put({"type": event_type, "data": data}),
            loop,
        )

    # Встаём в очередь
    with _queue_counter_lock:
        _queue_counter += 1
        position = _queue_counter

    # Если не первый — сообщаем позицию в очереди
    if position > 1:
        progress_callback("queue", {"position": position - 1})
        print(f"[QUEUE] Сессия {session_id} встала в очередь, позиция {position - 1}")

    try:
        # Ждём своей очереди (блокирующий вызов)
        with _audit_lock:
            # Сообщаем что очередь дошла
            progress_callback("queue", {"position": 0})

            try:
                # Определяем тип движка
                engine_type = _detect_engine(doc_type)

                if engine_type in SPECIAL_ENGINES:
                    # Спецдвижок (kpsc, drivers, kartochka_proekta)
                    progress_callback("audit_start", {
                        "doc_type": doc_type,
                        "filename": Path(target_path).name,
                        "engine": engine_type,
                    })

                    # Создаём session_dir для спецдвижка
                    from datetime import datetime as dt
                    timestamp = dt.now().strftime("%Y%m%d_%H%M%S")
                    session_dir = str(
                        Path(__file__).parent / "logs_result" / doc_type / f"session_{timestamp}"
                    )

                    result = _run_special_engine(engine_type, doc_type, target_path, session_dir)
                else:
                    # Стандартный pipeline (Vision/Paddle + rules.json)
                    from audit_engine.engine import AuditEngine

                    engine = AuditEngine(doc_type)
                    result = engine.run(
                        target_path,
                        progress_callback=progress_callback,
                    )

                session["result"] = {
                    "violations": result.violations,
                    "rules_checked": result.rules_checked,
                    "duration_sec": result.duration_sec,
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
                progress_callback("error", {"message": err_msg})

    finally:
        with _queue_counter_lock:
            _queue_counter -= 1
