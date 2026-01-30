from __future__ import annotations

import json
import os
import subprocess
import threading
import uuid
import zipfile
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, Optional

from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.responses import FileResponse

app = FastAPI(title="KPSC Validation Service")


def _runs_dir() -> Path:
    return Path(os.environ.get("RUNS_DIR", "/data/runs/kpsc_validation")).resolve()


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


def _parser_script(name: str) -> Path:
    return Path(__file__).resolve().parents[1] / "kpsc_validation" / "parser_scripts" / name


def _validator_scripts_dir() -> Path:
    return Path(__file__).resolve().parents[1] / "kpsc_validation" / "validation_scripts"


def _default_rules_path() -> Path:
    return Path(__file__).resolve().parents[1] / "kpsc_validation" / "validation_rules.json"


def _run_parsers(input_xlsx: Path, parser_outputs: Path) -> None:
    parser_outputs.mkdir(parents=True, exist_ok=True)

    tasks = [
        ("parse_kpsc_header.py", "kpsc_header_v2.json", ["-s", "КПСЦ"]),
        ("parse_kpsc_table1.py", "kpsc_table1_v2.json", ["-s", "КПСЦ"]),
        ("parse_legend.py", "legend_v2.json", ["-s", "Условные обозначения"]),
        ("parse_loss_digitization.py", "ocifrovka_poteri_v2.json", ["-s", "Оцифровка потерь КПСЦ"]),
        ("parse_pa1_chart.py", "pa1_chart_v3.json", ["-s", "ПА-1"]),
        ("parse_pa1_table.py", "pa1_table_v1.json", ["-s", "ПА-1"]),
        ("parse_pokazateli.py", "pokazateli_v3.json", ["-s", "Показатели"]),
        ("parse_spaghetti_sheet.py", "spaghetti_sheet_v2.json", ["-s", "Диаграмма Спагетти"]),
        ("parse_spaghetti_problems.py", "spaghetti_problems_v1.json", ["-s", "Перечень проблем по спагетти"]),
    ]

    for script_name, out_name, extra_args in tasks:
        script_path = _parser_script(script_name)
        out_path = parser_outputs / out_name
        cmd = ["python", str(script_path), "-i", str(input_xlsx), "-o", str(out_path), *extra_args]
        res = subprocess.run(cmd, capture_output=True, text=True)
        if res.returncode != 0:
            raise RuntimeError(f"Parser failed: {script_name}. STDERR: {res.stderr[:500]}")


def _run_validations(
    parser_outputs: Path,
    output_dir: Path,
    report_path: Path,
    rules_path: Path,
    parallel: int,
    verbose: bool,
) -> None:
    run_all = Path(__file__).resolve().parents[1] / "run_all_validations.py"

    cmd = [
        "python",
        str(run_all),
        "--parser-outputs",
        str(parser_outputs),
        "--output-dir",
        str(output_dir),
        "--report",
        str(report_path),
        "--rules",
        str(rules_path),
        "--scripts-dir",
        str(_validator_scripts_dir()),
        "--parallel",
        str(parallel),
    ]
    if verbose:
        cmd.append("--verbose")

    env = os.environ.copy()
    env["VALIDATION_RULES_PATH"] = str(rules_path)

    res = subprocess.run(cmd, capture_output=True, text=True, env=env)
    if res.returncode != 0:
        raise RuntimeError(f"Validation failed. STDERR: {res.stderr[:800]}")


def _run_job(
    job_id: str,
    input_path: Path,
    rules_path: Path,
    parallel: int,
    verbose: bool,
) -> None:
    run_dir = _job_dir(job_id)
    try:
        parser_outputs = run_dir / "parser_outputs"
        validation_outputs = run_dir / "validation_outputs"
        validation_report = run_dir / "validation_report.xlsx"

        _write_job_meta(
            job_id,
            {
                "job_id": job_id,
                "service": "kpsc_validation",
                "status": "processing",
                "created_at": datetime.utcnow().isoformat() + "Z",
                "rules_path": str(rules_path),
            },
        )

        _run_parsers(input_path, parser_outputs)
        _run_validations(
            parser_outputs=parser_outputs,
            output_dir=validation_outputs,
            report_path=validation_report,
            rules_path=rules_path,
            parallel=parallel,
            verbose=verbose,
        )

        _zip_run_dir(job_id)
        _write_job_meta(
            job_id,
            {
                **_read_job_meta(job_id),
                "status": "complete",
                "completed_at": datetime.utcnow().isoformat() + "Z",
                "artifacts": _list_artifacts(job_id),
            },
        )
    except Exception as exc:  # noqa: BLE001
        _write_job_meta(
            job_id,
            {
                "job_id": job_id,
                "service": "kpsc_validation",
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
    rules_file: Optional[UploadFile] = File(default=None),
    parallel: int = Form(default=5),
    verbose: bool = Form(default=False),
) -> Dict[str, Any]:
    if not file.filename or not file.filename.lower().endswith(".xlsx"):
        raise HTTPException(status_code=400, detail="Only .xlsx is supported")

    job_id = uuid.uuid4().hex
    run_dir = _job_dir(job_id)
    run_dir.mkdir(parents=True, exist_ok=True)

    input_path = run_dir / f"input_{Path(file.filename).name}"
    input_path.write_bytes(await file.read())

    if rules_file is not None:
        if not rules_file.filename or not rules_file.filename.lower().endswith(".json"):
            raise HTTPException(status_code=400, detail="rules_file must be .json")
        rules_path = run_dir / "validation_rules.json"
        rules_path.write_bytes(await rules_file.read())
    else:
        rules_path = _default_rules_path()

    _write_job_meta(
        job_id,
        {
            "job_id": job_id,
            "service": "kpsc_validation",
            "status": "queued",
            "created_at": datetime.utcnow().isoformat() + "Z",
            "input_filename": file.filename,
        },
    )

    t = threading.Thread(target=_run_job, args=(job_id, input_path, rules_path, int(parallel), bool(verbose)), daemon=True)
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
