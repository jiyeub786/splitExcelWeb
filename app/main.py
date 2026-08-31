import asyncio
import json
import queue
from io import BytesIO
from pathlib import Path
from typing import Optional

from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.responses import FileResponse, StreamingResponse
from fastapi.staticfiles import StaticFiles

from . import file_cache
from . import preview as preview_utils
from .job_manager import LOG_END_SENTINEL, JobStatus, job_manager
from .models import JobConfig

BASE_DIR = Path(__file__).resolve().parent.parent
STATIC_DIR = BASE_DIR / "static"

app = FastAPI(title="SplitExcel Web")


async def _resolve_bytes(file: Optional[UploadFile], token: Optional[str]) -> bytes:
    """미리보기 요청은 file(직접 업로드) 또는 token(/api/preview/upload로 미리 캐시한 파일)
    중 하나를 받는다. token을 쓰면 같은 파일을 여러 번 재전송하지 않아도 된다."""
    if token:
        data = file_cache.get(token)
        if data is None:
            raise HTTPException(
                status_code=400,
                detail="캐시된 파일을 찾을 수 없습니다(만료되었을 수 있음). 파일을 다시 선택해주세요.",
            )
        return data
    if file is not None:
        return await file.read()
    raise HTTPException(status_code=400, detail="file 또는 token 중 하나가 필요합니다")


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


@app.get("/api/jobs")
async def list_jobs():
    """Job 히스토리(최신순). 서버가 살아있는 동안의 기록만 대상이며 재시작하면 초기화된다."""
    return {"jobs": job_manager.list_jobs()}


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
        "queue_position": job_manager.get_queue_position(job_id),
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


@app.get("/api/jobs/{job_id}/results")
async def list_result_files(job_id: str):
    """zip 전체 대신 개별 결과 파일만 받고 싶을 때 목록을 보여주기 위함."""
    files = job_manager.list_result_files(job_id)
    if files is None:
        raise HTTPException(status_code=404, detail="작업을 찾을 수 없습니다")
    return {"files": files}


@app.get("/api/jobs/{job_id}/result/{filename}")
async def download_result_file(job_id: str, filename: str):
    job = job_manager.get(job_id)
    if job is None:
        raise HTTPException(status_code=404, detail="작업을 찾을 수 없습니다")
    # Path(...).name으로 경로 조작(디렉터리 탈출) 문자를 제거하고, 실제 이 Job의 결과 폴더
    # 안에 있는 파일인지 재확인한다.
    safe_name = Path(filename).name
    file_path = job.result_dir / safe_name
    if not file_path.is_file() or file_path.parent != job.result_dir:
        raise HTTPException(status_code=404, detail="결과 파일을 찾을 수 없습니다")
    return FileResponse(file_path, filename=safe_name)


@app.get("/api/system/excel-status")
async def excel_status():
    """작업 시작 전 '열려 있는 Excel 감지' 경고에 사용(읽기 전용, 종료하지 않음)."""
    return {"count": job_manager.count_excel_processes()}


@app.post("/api/preview/upload")
async def preview_upload(file: UploadFile = File(...)):
    """미리보기용 파일을 서버에 잠깐(최대 30분) 캐시해두고 토큰을 발급한다. 이후
    /api/preview/*, /api/validate/dry-run 요청은 파일을 재첨부하는 대신 이 토큰만 보내면 된다."""
    data = await file.read()
    token = file_cache.put(data)
    return {"token": token}


@app.post("/api/preview/sheets")
async def preview_sheets(
    file: Optional[UploadFile] = File(None), token: Optional[str] = Form(None)
):
    """시트명 자동완성 및 범위 선택 뷰어에서 사용. Excel COM 없이 즉시 응답."""
    data = await _resolve_bytes(file, token)
    try:
        sheets = preview_utils.list_sheet_names(BytesIO(data))
    except Exception as e:
        raise HTTPException(status_code=400, detail=f"엑셀 파일을 읽을 수 없습니다: {e}")
    return {"sheets": sheets}


@app.post("/api/preview/grid")
async def preview_grid(
    file: Optional[UploadFile] = File(None),
    token: Optional[str] = Form(None),
    sheet_name: str = Form(...),
    max_rows: int = Form(200),
    max_cols: int = Form(40),
):
    """범위 선택 뷰어의 그리드 미리보기(읽기 전용)."""
    data = await _resolve_bytes(file, token)
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
    file: Optional[UploadFile] = File(None),
    token: Optional[str] = Form(None),
    sheet_name: str = Form(...),
    column: str = Form(...),
    header_rows: int = Form(1),
):
    """분리 대상 목록 자동 추출: 지정한 열의 고유값을 등장 순서대로 반환."""
    data = await _resolve_bytes(file, token)
    try:
        values = preview_utils.extract_column_values(
            BytesIO(data), sheet_name, column, header_rows=header_rows
        )
    except ValueError as e:
        raise HTTPException(status_code=400, detail=str(e))
    except Exception as e:
        raise HTTPException(status_code=400, detail=f"엑셀 파일을 읽을 수 없습니다: {e}")
    return {"values": values, "count": len(values)}


@app.post("/api/preview/validate-split-list")
async def preview_validate_split_list(
    file: Optional[UploadFile] = File(None),
    token: Optional[str] = Form(None),
    sheet_name: str = Form(...),
    column: str = Form(...),
    header_rows: int = Form(1),
    split_list: str = Form(..., description="분리 대상 목록을 JSON 배열 문자열로 직렬화한 값"),
):
    """분리 대상 목록의 각 값이 실제로 필터 열에 존재하는지 사전 확인(오타 방지)."""
    try:
        values = json.loads(split_list)
        if not isinstance(values, list):
            raise ValueError("split_list는 배열이어야 합니다")
    except (json.JSONDecodeError, ValueError) as e:
        raise HTTPException(status_code=400, detail=f"split_list가 올바르지 않습니다: {e}")

    data = await _resolve_bytes(file, token)
    try:
        result = preview_utils.validate_split_list(
            BytesIO(data), sheet_name, column, header_rows, values
        )
    except ValueError as e:
        raise HTTPException(status_code=400, detail=str(e))
    except Exception as e:
        raise HTTPException(status_code=400, detail=f"엑셀 파일을 읽을 수 없습니다: {e}")
    return result


@app.post("/api/validate/dry-run")
async def validate_dry_run(
    source_file: Optional[UploadFile] = File(None),
    template_file: Optional[UploadFile] = File(None),
    source_token: Optional[str] = Form(None),
    template_token: Optional[str] = Form(None),
    sheet_tasks: str = Form(..., description="sheet_tasks를 JSON 배열 문자열로 직렬화한 값"),
):
    """작업 시작 전 사전 점검: 시트명이 원본/양식 양쪽에 모두 있는지, 범위 문법이 올바른지 확인해
    Excel COM을 띄우기 전에 흔한 실수를 미리 잡아낸다."""
    try:
        tasks = json.loads(sheet_tasks)
        if not isinstance(tasks, list):
            raise ValueError("sheet_tasks는 배열이어야 합니다")
    except (json.JSONDecodeError, ValueError) as e:
        raise HTTPException(status_code=400, detail=f"sheet_tasks가 올바르지 않습니다: {e}")

    source_data = await _resolve_bytes(source_file, source_token)
    template_data = await _resolve_bytes(template_file, template_token)
    try:
        errors = preview_utils.dry_run_validate(BytesIO(source_data), BytesIO(template_data), tasks)
    except Exception as e:
        raise HTTPException(status_code=400, detail=f"엑셀 파일을 읽을 수 없습니다: {e}")
    return {"ok": len(errors) == 0, "errors": errors}


app.mount("/", StaticFiles(directory=str(STATIC_DIR), html=True), name="static")
