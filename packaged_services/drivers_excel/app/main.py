from __future__ import annotations

import json
import os
import threading
import uuid
import zipfile
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, Optional

import yaml
from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.responses import FileResponse
from openai import OpenAI

from drivers.analyzer import analyze_sections, collect_remarks_and_summaries, export_missing_driver_report
from drivers.parser import parse_excel_to_json

app = FastAPI(title="Drivers Excel Service")


def _runs_dir() -> Path:
    return Path(os.environ.get("RUNS_DIR", "/data/runs/drivers_excel")).resolve()


def _job_dir(job_id: str) -> Path:
    return _runs_dir() / job_id


def _job_meta_path(job_id: str) -> Path:
    return _job_dir(job_id) / "job.json"


def _write_job_meta(job_id: str, payload: Dict[str, Any]) -> None:
    path = _job_meta_path(job_id)
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")


def _read_job_meta(job_id: str) -> Dict[str, Any]:
    path = _job_meta_path(job_id)
    if not path.exists():
        raise FileNotFoundError(job_id)
    return json.loads(path.read_text(encoding="utf-8"))


def _list_artifacts(job_id: str) -> list[str]:
    root = _job_dir(job_id)
    if not root.exists():
        raise FileNotFoundError(job_id)
    return sorted([p.name for p in root.iterdir() if p.is_file()])


def _safe_artifact_path(job_id: str, name: str) -> Path:
    safe_name = Path(name).name
    path = _job_dir(job_id) / safe_name
    if not path.exists():
        raise FileNotFoundError(name)
    return path


def _load_config() -> Dict[str, Any]:
    config_path = Path(__file__).resolve().parents[1] / "config" / "config.yaml"
    return yaml.safe_load(config_path.read_text(encoding="utf-8"))


def _get_openai_client() -> OpenAI:
    api_key = os.environ.get("OPENAI_API_KEY")
    if not api_key:
        raise RuntimeError("OPENAI_API_KEY is not set")
    return OpenAI(api_key=api_key)


def _zip_run_dir(job_id: str, zip_name: str = "artifacts.zip") -> Path:
    root = _job_dir(job_id)
    zip_path = root / zip_name
    with zipfile.ZipFile(zip_path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        for p in sorted(root.iterdir()):
            if p.is_file() and p.name != zip_name:
                zf.write(p, arcname=p.name)
    return zip_path


def _run_job(
    job_id: str,
    input_path: Path,
    original_filename: str,
    primary_threshold: Optional[float],
    fallback_threshold: Optional[float],
    model: Optional[str],
    temperature: Optional[float],
    sheet_name: Optional[str],
) -> None:
    run_dir = _job_dir(job_id)
    try:
        cfg = _load_config()
        analysis_cfg = cfg.get("analysis", {})
        openai_cfg = cfg.get("openai", {})
        output_cfg = cfg.get("output", {})

        primary = float(primary_threshold) if primary_threshold is not None else float(analysis_cfg.get("primary_threshold", 9.0))
        fallback = float(fallback_threshold) if fallback_threshold is not None else float(analysis_cfg.get("fallback_threshold", 7.0))
        model_name = model or os.environ.get("DRIVERS_MODEL") or openai_cfg.get("model", "gpt-4.1-mini-2025-04-14")
        temp = float(temperature) if temperature is not None else float(openai_cfg.get("temperature", 0))
        sheet = sheet_name or output_cfg.get("sheet_name", "Итог")

        _write_job_meta(
            job_id,
            {
                "job_id": job_id,
                "service": "drivers_excel",
                "status": "processing",
                "created_at": datetime.utcnow().isoformat() + "Z",
                "input_filename": original_filename,
                "config": {
                    "primary_threshold": primary,
                    "fallback_threshold": fallback,
                    "model": model_name,
                    "temperature": temp,
                    "sheet_name": sheet,
                },
            },
        )

        parsed = parse_excel_to_json(input_path)
        (run_dir / "01_parsed.json").write_text(json.dumps(parsed, ensure_ascii=False, indent=2), encoding="utf-8")

        client = _get_openai_client()
        section_results = analyze_sections(
            parsed_data=parsed,
            client=client,
            primary_threshold=primary,
            fallback_threshold=fallback,
            model=model_name,
            temperature=temp,
        )

        remarks, driver_summary_map = collect_remarks_and_summaries(section_results)
        bundle = {"section_results": section_results, "remarks": remarks, "driver_summary_map": driver_summary_map}
        (run_dir / "02_llm_analysis.json").write_text(json.dumps(bundle, ensure_ascii=False, indent=2), encoding="utf-8")

        report_path = run_dir / "03_report.xlsx"
        export_missing_driver_report(parsed_data=parsed, section_results=section_results, output_path=report_path, sheet_name=sheet)

        _zip_run_dir(job_id)
        _write_job_meta(
            job_id,
            {
                **_read_job_meta(job_id),
                "status": "complete",
                "completed_at": datetime.utcnow().isoformat() + "Z",
                "remarks_count": len(remarks),
                "artifacts": _list_artifacts(job_id),
            },
        )
    except Exception as exc:  # noqa: BLE001
        _write_job_meta(
            job_id,
            {
                "job_id": job_id,
                "service": "drivers_excel",
                "status": "error",
                "error": str(exc),
                "artifacts": _list_artifacts(job_id) if run_dir.exists() else [],
            },
        )


@app.get("/health")
def health() -> Dict[str, str]:
    return {"status": "ok"}


@app.post("/jobs")
async def create_job(
    file: UploadFile = File(...),
    primary_threshold: Optional[float] = Form(default=None),
    fallback_threshold: Optional[float] = Form(default=None),
    model: Optional[str] = Form(default=None),
    temperature: Optional[float] = Form(default=None),
    sheet_name: Optional[str] = Form(default=None),
) -> Dict[str, Any]:
    if not file.filename or not file.filename.lower().endswith(".xlsx"):
        raise HTTPException(status_code=400, detail="Only .xlsx is supported")

    job_id = uuid.uuid4().hex
    run_dir = _job_dir(job_id)
    run_dir.mkdir(parents=True, exist_ok=True)

    input_path = run_dir / f"00_input_{Path(file.filename).name}"
    content = await file.read()
    input_path.write_bytes(content)

    _write_job_meta(
        job_id,
        {
            "job_id": job_id,
            "service": "drivers_excel",
            "status": "queued",
            "created_at": datetime.utcnow().isoformat() + "Z",
            "input_filename": file.filename,
        },
    )

    t = threading.Thread(
        target=_run_job,
        args=(job_id, input_path, file.filename, primary_threshold, fallback_threshold, model, temperature, sheet_name),
        daemon=True,
    )
    t.start()

    return {"job_id": job_id}


@app.get("/jobs/{job_id}")
def get_job(job_id: str) -> Dict[str, Any]:
    try:
        return _read_job_meta(job_id)
    except FileNotFoundError:
        raise HTTPException(status_code=404, detail="job not found")


@app.get("/jobs/{job_id}/artifacts")
def list_job_artifacts(job_id: str) -> Dict[str, Any]:
    try:
        return {"job_id": job_id, "artifacts": _list_artifacts(job_id)}
    except FileNotFoundError:
        raise HTTPException(status_code=404, detail="job not found")


@app.get("/jobs/{job_id}/artifacts/{name}")
def download_artifact(job_id: str, name: str) -> FileResponse:
    try:
        path = _safe_artifact_path(job_id, name)
    except FileNotFoundError:
        raise HTTPException(status_code=404, detail="artifact not found")
    return FileResponse(path)
