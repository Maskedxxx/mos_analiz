from __future__ import annotations

import json
import os
import threading
import uuid
import zipfile
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, Optional

from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.responses import FileResponse

from word_audit.audit import audit_order, audit_package, audit_reg, write_outputs

app = FastAPI(title="Word Audit Service")


def _runs_dir() -> Path:
    return Path(os.environ.get("RUNS_DIR", "/data/runs/word_audit")).resolve()


def _job_dir(job_id: str) -> Path:
    return _runs_dir() / job_id


def _job_meta_path(job_id: str) -> Path:
    return _job_dir(job_id) / "job.json"


def _write_job_meta(job_id: str, payload: Dict[str, Any]) -> None:
    _job_meta_path(job_id).write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")


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


def _zip_run_dir(job_id: str, zip_name: str = "artifacts.zip") -> Path:
    root = _job_dir(job_id)
    zip_path = root / zip_name
    with zipfile.ZipFile(zip_path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        for p in sorted(root.iterdir()):
            if p.is_file() and p.name != zip_name:
                zf.write(p, arcname=p.name)
    return zip_path


def _load_ruleset(kind: str) -> Dict[str, Any]:
    mapping_path = Path(__file__).resolve().parents[1] / "rules" / "ruleset_mapping.json"
    mapping = json.loads(mapping_path.read_text(encoding="utf-8"))
    return mapping[kind]


def _run_job(
    job_id: str,
    job_type: str,
    model: str,
    temperature: float,
    paths: Dict[str, Path],
) -> None:
    run_dir = _job_dir(job_id)
    try:
        _write_job_meta(
            job_id,
            {
                **_read_job_meta(job_id),
                "status": "processing",
                "started_at": datetime.utcnow().isoformat() + "Z",
            },
        )
        if job_type in ("order_comp_ppu", "order_ppu"):
            ruleset = _load_ruleset("order")
            violations, debug = audit_order(
                target_path=paths["target"],
                template_path=paths["template"],
                tz_path=paths["tz"],
                model=model,
                temperature=temperature,
                ruleset=ruleset,
            )
        elif job_type in ("reg_comp_ppu", "reg_ppu"):
            ruleset = _load_ruleset(job_type)
            violations, debug = audit_reg(
                target_path=paths["target"],
                template_path=paths["template"],
                tz_path=paths["tz"],
                model=model,
                temperature=temperature,
                ruleset=ruleset,
                reg_type=job_type,
            )
        elif job_type == "package":
            violations, debug = audit_package(
                order_comp_path=paths["order_comp"],
                reg_comp_path=paths["reg_comp"],
                order_ppu_path=paths["order_ppu"],
                reg_ppu_path=paths["reg_ppu"],
                model=model,
                temperature=temperature,
            )
        else:
            raise ValueError(f"Unsupported job_type: {job_type}")

        out = write_outputs(run_dir, violations, debug)
        _zip_run_dir(job_id)
        _write_job_meta(
            job_id,
            {
                **_read_job_meta(job_id),
                "status": "complete",
                "completed_at": datetime.utcnow().isoformat() + "Z",
                "violations_count": len(violations),
                "outputs": out,
                "artifacts": _list_artifacts(job_id),
            },
        )
    except Exception as exc:  # noqa: BLE001
        _write_job_meta(
            job_id,
            {
                "job_id": job_id,
                "service": "word_audit",
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
    job_type: str = Form(...),
    model: Optional[str] = Form(default=None),
    temperature: float = Form(default=0.0),
    target_doc: Optional[UploadFile] = File(default=None),
    template_doc: Optional[UploadFile] = File(default=None),
    tz_doc: Optional[UploadFile] = File(default=None),
    order_comp_doc: Optional[UploadFile] = File(default=None),
    reg_comp_doc: Optional[UploadFile] = File(default=None),
    order_ppu_doc: Optional[UploadFile] = File(default=None),
    reg_ppu_doc: Optional[UploadFile] = File(default=None),
) -> Dict[str, Any]:
    model_name = model or os.environ.get("WORD_AUDIT_MODEL", "gpt-4.1-2025-04-14")
    job_id = uuid.uuid4().hex
    run_dir = _job_dir(job_id)
    run_dir.mkdir(parents=True, exist_ok=True)

    paths: Dict[str, Path] = {}

    async def save_upload(key: str, up: UploadFile) -> None:
        if not up.filename or not up.filename.lower().endswith(".docx"):
            raise HTTPException(status_code=400, detail=f"{key} must be .docx")
        path = run_dir / f"{key}_{Path(up.filename).name}"
        paths[key] = path
        path.write_bytes(await up.read())

    if job_type == "package":
        missing = [k for k, v in {"order_comp": order_comp_doc, "reg_comp": reg_comp_doc, "order_ppu": order_ppu_doc, "reg_ppu": reg_ppu_doc}.items() if v is None]
        if missing:
            raise HTTPException(status_code=400, detail=f"Missing files for package: {', '.join(missing)}")
        await save_upload("order_comp", order_comp_doc)  # type: ignore[arg-type]
        await save_upload("reg_comp", reg_comp_doc)      # type: ignore[arg-type]
        await save_upload("order_ppu", order_ppu_doc)    # type: ignore[arg-type]
        await save_upload("reg_ppu", reg_ppu_doc)        # type: ignore[arg-type]
    else:
        if target_doc is None or template_doc is None or tz_doc is None:
            raise HTTPException(status_code=400, detail="target_doc, template_doc, tz_doc are required")
        await save_upload("target", target_doc)
        await save_upload("template", template_doc)
        await save_upload("tz", tz_doc)

    _write_job_meta(
        job_id,
        {
            "job_id": job_id,
            "service": "word_audit",
            "status": "queued",
            "created_at": datetime.utcnow().isoformat() + "Z",
            "job_type": job_type,
            "model": model_name,
            "temperature": float(temperature),
        },
    )

    t = threading.Thread(target=_run_job, args=(job_id, job_type, model_name, float(temperature), paths), daemon=True)
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
