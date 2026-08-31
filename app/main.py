import asyncio
import queue
from io import BytesIO
from pathlib import Path

from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.responses import FileResponse, StreamingResponse
from fastapi.staticfiles import StaticFiles

from . import preview as preview_utils
from .job_manager import LOG_END_SENTINEL, JobStatus, job_manager
from .models import JobConfig

BASE_DIR = Path(__file__).resolve().parent.parent
STATIC_DIR = BASE_DIR / "static"

app = FastAPI(title="SplitExcel Web")


@app.post("/api/jobs")
async def create_job(
    source_file: UploadFile = File(..., description="원본파일(splitExcel의 file01)"),
    template_file: UploadFile = File(..., description="양식파일(splitExcel의 file02)"),
    config: str = Form(..., description="JobConfig를 JSON 문자열로 직렬화한 값"),
):
    try:
        cfg = JobConfig.model_validate_json(config)
    except Exception as e:
        raise HTTPException(status_code=400, detail=f"설정값이 올바르지 않습니다: {e}")

    if not cfg.sheet_tasks:
        raise HTTPException(status_code=400, detail="시트 작업 정의(sheet_tasks)가 비어 있습니다")
    if not cfg.split_list:
        raise HTTPException(status_code=400, detail="분리 대상 목록(split_list)이 비어 있습니다")

    job = job_manager.create_job(cfg.model_dump(), source_file, template_file)
    return {"job_id": job.id, "status": job.status}


@app.get("/api/jobs/{job_id}")
async def get_job(job_id: str):
    job = job_manager.get(job_id)
    if job is None:
        raise HTTPException(status_code=404, detail="작업을 찾을 수 없습니다")
    return {
        "job_id": job.id,
        "status": job.status,
        "error": job.error,
        "created_at": job.created_at.isoformat(),
        "result_ready": job.status == JobStatus.DONE and job.result_zip_path.exists(),
    }


@app.post("/api/jobs/{job_id}/cancel")
async def cancel_job(job_id: str):
    """데스크톱 버전의 '강제 종료' 버튼(협조적 취소)에 대응."""
    if not job_manager.cancel(job_id):
        raise HTTPException(status_code=404, detail="작업을 찾을 수 없습니다")
    return {"ok": True}


@app.post("/api/jobs/{job_id}/force-stop-excel")
async def force_stop_excel(job_id: str):
    """
    데스크톱 버전에서 강제종료를 2번 눌렀을 때의 'excel.exe 프로세스 강제 종료'에 대응.
    작업이 단일 워커로 직렬 처리되므로 다른 사용자의 작업에 영향을 줄 위험이 낮지만,
    운영 환경에서 동시 접속자가 늘어나면 재검토가 필요하다 (MIGRATION_PLAN.md 5장 참고).
    """
    job = job_manager.get(job_id)
    if job is None:
        raise HTTPException(status_code=404, detail="작업을 찾을 수 없습니다")
    job.cancel_event.set()
    killed = job_manager.force_kill_excel()
    return {"ok": True, "killed_processes": killed}


@app.get("/api/jobs/{job_id}/logs")
async def stream_logs(job_id: str):
    """LogWindow(Qt Signal 기반 실시간 로그)의 웹 대응 — Server-Sent Events로 스트리밍."""
    job = job_manager.get(job_id)
    if job is None:
        raise HTTPException(status_code=404, detail="작업을 찾을 수 없습니다")

    async def event_generator():
        while True:
            try:
                line = job.log_queue.get_nowait()
            except queue.Empty:
                if job.finished:
                    break
                await asyncio.sleep(0.3)
                continue

            if line == LOG_END_SENTINEL:
                break
            yield f"data: {line}\n\n"

        yield f"event: done\ndata: {job.status.value}\n\n"

    return StreamingResponse(event_generator(), media_type="text/event-stream")


@app.get("/api/jobs/{job_id}/result")
async def download_result(job_id: str):
    job = job_manager.get(job_id)
    if job is None:
        raise HTTPException(status_code=404, detail="작업을 찾을 수 없습니다")
    if job.status != JobStatus.DONE:
        raise HTTPException(status_code=409, detail="아직 완료되지 않은 작업입니다")
    if not job.result_zip_path.exists():
        raise HTTPException(status_code=404, detail="결과 파일이 없습니다")
    return FileResponse(job.result_zip_path, filename=f"SplitExcel_{job_id}.zip")


@app.post("/api/preview/sheets")
async def preview_sheets(file: UploadFile = File(...)):
    """시트명 자동완성 및 범위 선택 뷰어에서 사용. Excel COM 없이 즉시 응답."""
    data = await file.read()
    try:
        sheets = preview_utils.list_sheet_names(BytesIO(data))
    except Exception as e:
        raise HTTPException(status_code=400, detail=f"엑셀 파일을 읽을 수 없습니다: {e}")
    return {"sheets": sheets}


@app.post("/api/preview/grid")
async def preview_grid(
    file: UploadFile = File(...),
    sheet_name: str = Form(...),
    max_rows: int = Form(200),
    max_cols: int = Form(40),
):
    """범위 선택 뷰어의 그리드 미리보기(읽기 전용)."""
    data = await file.read()
    try:
        grid = preview_utils.read_grid(
            BytesIO(data), sheet_name, max_rows=max_rows, max_cols=max_cols
        )
    except ValueError as e:
        raise HTTPException(status_code=400, detail=str(e))
    except Exception as e:
        raise HTTPException(status_code=400, detail=f"엑셀 파일을 읽을 수 없습니다: {e}")
    return grid


@app.post("/api/preview/column-values")
async def preview_column_values(
    file: UploadFile = File(...),
    sheet_name: str = Form(...),
    column: str = Form(...),
    header_rows: int = Form(1),
):
    """분리 대상 목록 자동 추출: 지정한 열의 고유값을 등장 순서대로 반환."""
    data = await file.read()
    try:
        values = preview_utils.extract_column_values(
            BytesIO(data), sheet_name, column, header_rows=header_rows
        )
    except ValueError as e:
        raise HTTPException(status_code=400, detail=str(e))
    except Exception as e:
        raise HTTPException(status_code=400, detail=f"엑셀 파일을 읽을 수 없습니다: {e}")
    return {"values": values, "count": len(values)}


app.mount("/", StaticFiles(directory=str(STATIC_DIR), html=True), name="static")
