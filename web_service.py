#!/usr/bin/env python3
"""Веб-сервис для анализа драйверов производительности."""

import asyncio
import json
import os
import shutil
import uuid
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime
from pathlib import Path
from typing import Dict

import yaml
from fastapi import FastAPI, File, UploadFile, WebSocket, WebSocketDisconnect, HTTPException
from fastapi.responses import FileResponse, HTMLResponse
from fastapi.templating import Jinja2Templates
from fastapi import Request
from openai import OpenAI

from analyzer import analyze_sections, collect_remarks_and_summaries, export_missing_driver_report
from parser import parse_excel_to_json

app = FastAPI(title="Driver Analysis Service")

# Настройка шаблонов
templates = Jinja2Templates(directory="templates")

# Хранилище активных сессий
sessions: Dict[str, Dict] = {}


class ConnectionManager:
    """Менеджер WebSocket соединений для отправки прогресса."""

    def __init__(self):
        self.active_connections: Dict[str, WebSocket] = {}

    async def connect(self, session_id: str, websocket: WebSocket):
        await websocket.accept()
        self.active_connections[session_id] = websocket

    def disconnect(self, session_id: str):
        if session_id in self.active_connections:
            del self.active_connections[session_id]

    async def send_progress(self, session_id: str, message: str, current: int, total: int):
        if session_id in self.active_connections:
            try:
                await self.active_connections[session_id].send_json(
                    {"status": "progress", "message": message, "current": current, "total": total}
                )
            except Exception:
                pass

    async def send_complete(self, session_id: str, remarks_count: int):
        if session_id in self.active_connections:
            try:
                await self.active_connections[session_id].send_json(
                    {"status": "complete", "remarks_count": remarks_count}
                )
            except Exception:
                pass

    async def send_error(self, session_id: str, error: str):
        if session_id in self.active_connections:
            try:
                await self.active_connections[session_id].send_json({"status": "error", "error": error})
            except Exception:
                pass


manager = ConnectionManager()


def load_config(config_path: Path = Path("config.yaml")) -> Dict:
    """Загружает конфигурацию из YAML файла."""
    with open(config_path, "r", encoding="utf-8") as f:
        return yaml.safe_load(f)


def get_openai_client(config: Dict) -> OpenAI:
    """Создает клиента OpenAI."""
    api_key = config.get("openai", {}).get("api_key")
    if not api_key:
        api_key = os.getenv("OPENAI_API_KEY")
    if not api_key:
        raise ValueError("OpenAI API key not found")
    return OpenAI(api_key=api_key)


def create_session_directory(results_dir: Path, input_filename: str) -> tuple[Path, str]:
    """Создает уникальную папку для сессии."""
    timestamp = datetime.now().strftime("%Y-%m-%d_%H%M%S")
    short_uuid = str(uuid.uuid4())[:8]
    session_name = f"{timestamp}_{short_uuid}"

    session_dir = results_dir / session_name
    session_dir.mkdir(parents=True, exist_ok=True)

    return session_dir, session_name


def process_file_sync(session_id: str, input_file: Path, config: Dict, loop: asyncio.AbstractEventLoop):
    """Синхронная обработка файла (запускается в отдельном потоке)."""
    # 1. Создание папки сессии
    results_dir = Path(config.get("output", {}).get("results_dir", "results"))
    session_dir, _ = create_session_directory(results_dir, input_file.name)
    sessions[session_id]["session_dir"] = session_dir

    # 2. Копирование входного файла
    input_copy = session_dir / f"00_input_{input_file.name}"
    shutil.copy2(input_file, input_copy)

    # 3. Парсинг Excel -> JSON
    asyncio.run_coroutine_threadsafe(
        manager.send_progress(session_id, "Парсинг Excel файла...", 2, 6), loop
    )

    parsed_data = parse_excel_to_json(input_file)
    parsed_json_path = session_dir / "01_parsed.json"
    parsed_json_path.write_text(json.dumps(parsed_data, ensure_ascii=False, indent=2), encoding="utf-8")

    asyncio.run_coroutine_threadsafe(
        manager.send_progress(session_id, "Запуск анализа LLM...", 3, 6), loop
    )

    # 4. Анализ с помощью LLM
    client = get_openai_client(config)

    analysis_config = config.get("analysis", {})
    primary_threshold = analysis_config.get("primary_threshold", 9.0)
    fallback_threshold = analysis_config.get("fallback_threshold", 7.0)

    openai_config = config.get("openai", {})
    model = openai_config.get("model", "gpt-4o-mini")
    temperature = openai_config.get("temperature", 0)

    # Callback для прогресса анализа секций
    def progress_callback(message: str, current: int, total: int):
        asyncio.run_coroutine_threadsafe(
            manager.send_progress(session_id, message, 3 + current, 6), loop
        )

    section_results = analyze_sections(
        parsed_data=parsed_data,
        client=client,
        primary_threshold=primary_threshold,
        fallback_threshold=fallback_threshold,
        model=model,
        temperature=temperature,
        progress_callback=progress_callback,
    )

    asyncio.run_coroutine_threadsafe(
        manager.send_progress(session_id, "Агрегация результатов...", 5, 6), loop
    )

    # 5. Агрегация результатов
    remarks, driver_summary_map = collect_remarks_and_summaries(section_results)

    result_bundle = {
        "section_results": section_results,
        "remarks": remarks,
        "driver_summary_map": driver_summary_map,
    }

    llm_json_path = session_dir / "02_llm_analysis.json"
    llm_json_path.write_text(json.dumps(result_bundle, ensure_ascii=False, indent=2), encoding="utf-8")

    asyncio.run_coroutine_threadsafe(
        manager.send_progress(session_id, "Генерация Excel отчета...", 6, 6), loop
    )

    # 6. Генерация Excel отчета
    output_excel = session_dir / "03_report.xlsx"
    sheet_name = config.get("output", {}).get("sheet_name", "Итог")

    export_missing_driver_report(
        parsed_data=parsed_data,
        section_results=section_results,
        output_path=output_excel,
        sheet_name=sheet_name,
    )

    # 7. Обновление сессии
    sessions[session_id]["status"] = "complete"
    sessions[session_id]["report_path"] = output_excel
    sessions[session_id]["remarks_count"] = len(remarks)

    asyncio.run_coroutine_threadsafe(manager.send_complete(session_id, len(remarks)), loop)

    return True


async def process_file_async(session_id: str, input_file: Path, config: Dict):
    """Асинхронная обёртка для обработки файла."""
    loop = asyncio.get_event_loop()
    await manager.send_progress(session_id, "Подготовка файла...", 1, 6)

    try:
        # Запуск синхронной обработки в отдельном потоке
        executor = ThreadPoolExecutor(max_workers=1)
        await loop.run_in_executor(executor, process_file_sync, session_id, input_file, config, loop)
    except Exception as e:
        sessions[session_id]["status"] = "error"
        sessions[session_id]["error"] = str(e)
        await manager.send_error(session_id, str(e))


@app.get("/", response_class=HTMLResponse)
async def read_root(request: Request):
    """Главная страница с интерфейсом загрузки."""
    return templates.TemplateResponse("index.html", {"request": request})


@app.post("/upload")
async def upload_file(file: UploadFile = File(...)):
    """Загрузка Excel файла и запуск обработки."""
    # Проверка формата
    if not file.filename.endswith(".xlsx"):
        raise HTTPException(status_code=400, detail="Только .xlsx файлы поддерживаются")

    # Создание сессии
    session_id = str(uuid.uuid4())
    sessions[session_id] = {"status": "processing", "filename": file.filename}

    # Сохранение временного файла
    temp_dir = Path("uploads")
    temp_dir.mkdir(exist_ok=True)
    temp_file = temp_dir / f"{session_id}_{file.filename}"

    with open(temp_file, "wb") as f:
        content = await file.read()
        f.write(content)

    # Загрузка конфигурации
    config = load_config()

    # Запуск обработки в фоне
    asyncio.create_task(process_file_async(session_id, temp_file, config))

    return {"session_id": session_id}


@app.websocket("/ws/{session_id}")
async def websocket_endpoint(websocket: WebSocket, session_id: str):
    """WebSocket для отслеживания прогресса обработки."""
    await manager.connect(session_id, websocket)
    try:
        while True:
            # Держим соединение открытым
            await websocket.receive_text()
    except WebSocketDisconnect:
        manager.disconnect(session_id)


@app.get("/download/{session_id}/report")
async def download_report(session_id: str):
    """Скачивание Excel отчета."""
    if session_id not in sessions:
        raise HTTPException(status_code=404, detail="Сессия не найдена")

    session = sessions[session_id]
    if session["status"] != "complete":
        raise HTTPException(status_code=400, detail="Обработка еще не завершена")

    report_path = session["report_path"]
    if not report_path.exists():
        raise HTTPException(status_code=404, detail="Отчет не найден")

    return FileResponse(
        report_path,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename=f"report_{session_id[:8]}.xlsx",
    )


@app.get("/download/{session_id}/json")
async def download_json(session_id: str):
    """Скачивание JSON результатов."""
    if session_id not in sessions:
        raise HTTPException(status_code=404, detail="Сессия не найдена")

    session = sessions[session_id]
    if session["status"] != "complete":
        raise HTTPException(status_code=400, detail="Обработка еще не завершена")

    json_path = session["session_dir"] / "02_llm_analysis.json"
    if not json_path.exists():
        raise HTTPException(status_code=404, detail="JSON не найден")

    return FileResponse(json_path, media_type="application/json", filename=f"analysis_{session_id[:8]}.json")


@app.get("/status/{session_id}")
async def get_status(session_id: str):
    """Получение текущего статуса сессии."""
    if session_id not in sessions:
        raise HTTPException(status_code=404, detail="Сессия не найдена")

    return sessions[session_id]


if __name__ == "__main__":
    import uvicorn

    print("🚀 Запуск веб-сервиса на http://localhost:8000")
    print("📁 Откройте браузер: http://localhost:8000")
    uvicorn.run(app, host="0.0.0.0", port=8000)
