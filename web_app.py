from __future__ import annotations

import tempfile
import uuid
from pathlib import Path

from fastapi import FastAPI, File, HTTPException, UploadFile
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles
from starlette.background import BackgroundTask

from mb_calculator import build_output


BASE_DIR = Path(__file__).resolve().parent
STATIC_DIR = BASE_DIR / "web_static"

app = FastAPI(title="MB Calculator")
app.mount("/static", StaticFiles(directory=STATIC_DIR), name="static")


def cleanup_files(*paths: Path) -> None:
    for path in paths:
        try:
            path.unlink(missing_ok=True)
        except OSError:
            pass


@app.get("/")
def index() -> FileResponse:
    return FileResponse(STATIC_DIR / "index.html")


@app.get("/health")
def health() -> dict[str, str]:
    return {"status": "ok"}


@app.post("/calculate")
async def calculate(file: UploadFile = File(...)) -> FileResponse:
    filename = file.filename or "input.xlsx"
    if not filename.lower().endswith((".xlsx", ".xlsm")):
        raise HTTPException(status_code=400, detail="Please upload an .xlsx or .xlsm file.")

    work_dir = Path(tempfile.gettempdir()) / "mb_calculator_uploads"
    work_dir.mkdir(parents=True, exist_ok=True)

    token = uuid.uuid4().hex
    input_path = work_dir / f"{token}_{Path(filename).stem}.xlsx"
    output_path = work_dir / f"{token}_calculated_output.xlsx"

    content = await file.read()
    if not content:
        raise HTTPException(status_code=400, detail="Uploaded file is empty.")

    input_path.write_bytes(content)

    try:
        build_output(input_path, output_path)
    except KeyError as exc:
        cleanup_files(input_path, output_path)
        raise HTTPException(
            status_code=400,
            detail=f"Workbook is missing the expected sheet: {exc}",
        ) from exc
    except Exception as exc:
        cleanup_files(input_path, output_path)
        raise HTTPException(
            status_code=500,
            detail=f"Could not calculate the workbook: {exc}",
        ) from exc

    return FileResponse(
        output_path,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename=f"{Path(filename).stem}_calculated.xlsx",
        background=BackgroundTask(cleanup_files, input_path, output_path),
    )
