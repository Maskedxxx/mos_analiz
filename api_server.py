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
import json
import os
import shutil
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

# CORS для dev-режима (Vite на :5173)
app.add_middleware(
    CORSMiddleware,
    allow_origins=["http://localhost:5173"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# --- Конфигурация ---
AUTH_LOGIN = os.getenv("AUDIT_LOGIN", "admin")
AUTH_PASSWORD = os.getenv("AUDIT_PASSWORD", "admin")
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

    xlsx_path = Path(session["session_dir"]) / "audit_result.xlsx"
    if not xlsx_path.exists():
        raise HTTPException(status_code=404, detail="Excel не сформирован")

    return FileResponse(
        str(xlsx_path),
        filename=f"audit_{session_id}.xlsx",
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


# ── Фоновый поток аудита ─────────────────────────────────────────

def _run_audit_thread(session_id: str, doc_type: str, target_path: str):
    """Запуск AuditEngine.run() в фоновом потоке с progress_callback."""
    session = sessions[session_id]
    loop = session["loop"]
    queue = session["events"]

    def progress_callback(event_type: str, data: dict):
        """Колбэк из engine — пушит события в SSE-очередь."""
        asyncio.run_coroutine_threadsafe(
            queue.put({"type": event_type, "data": data}),
            loop,
        )

    try:
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
        session["status"] = "error"
        progress_callback("error", {"message": str(e)})
