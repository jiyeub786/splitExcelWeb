"""
Job(작업) 생명주기 관리.

MIGRATION_PLAN.md에서 정리한 위험요소를 반영한 설계:
- Excel COM은 동시 다중 자동화에 취약하므로, 작업은 단일 워커 스레드가 큐에서 하나씩 꺼내
  "직렬로" 처리한다 (Phase 5: 동시성 대응).
- 기존 데스크톱 버전의 전역 flag.txt 취소 로직은 Job 단위 threading.Event로 대체해 여러 Job이
  섞이지 않도록 격리한다 (Phase 2 체크리스트).
- 기존 LogWindow의 커스텀 logging.Handler 아이디어를 그대로 가져오되, 대상만 Qt Signal이 아니라
  Job별 Queue(→ SSE)로 바꿨다.
"""

import datetime
import logging
import queue
import shutil
import threading
import uuid
import zipfile
from dataclasses import dataclass, field
from enum import Enum
from pathlib import Path
from typing import Dict, List, Optional

import pythoncom

from .excel_processor import ExcelSplitProcessor

BASE_DIR = Path(__file__).resolve().parent.parent
STORAGE_DIR = BASE_DIR / "storage"
UPLOAD_DIR = STORAGE_DIR / "uploads"
RESULT_DIR = STORAGE_DIR / "results"

for _d in (UPLOAD_DIR, RESULT_DIR):
    _d.mkdir(parents=True, exist_ok=True)

LOG_END_SENTINEL = "__END__"
JOB_RETENTION_DAYS = 7  # storage/uploads, storage/results 자동 정리 보관 기간


class JobStatus(str, Enum):
    PENDING = "pending"
    RUNNING = "running"
    DONE = "done"
    ERROR = "error"
    CANCELLED = "cancelled"


class QueueLogHandler(logging.Handler):
    """logging 레코드를 Job 전용 큐에 밀어넣는다. LogWindow의 emit_log 오버라이드를 대체."""

    def __init__(self, log_queue: "queue.Queue[str]"):
        super().__init__()
        self.log_queue = log_queue

    def emit(self, record: logging.LogRecord) -> None:
        try:
            self.log_queue.put(self.format(record))
        except Exception:
            pass


@dataclass
class Job:
    id: str
    config: dict
    source_path: Path
    template_path: Path
    result_dir: Path
    status: JobStatus = JobStatus.PENDING
    error: Optional[str] = None
    created_at: datetime.datetime = field(default_factory=datetime.datetime.now)
    cancel_event: threading.Event = field(default_factory=threading.Event)
    log_queue: "queue.Queue[str]" = field(default_factory=queue.Queue)
    finished: bool = False

    @property
    def result_zip_path(self) -> Path:
        return self.result_dir.with_suffix(".zip")


class JobManager:
    def __init__(self) -> None:
        self.jobs: Dict[str, Job] = {}
        self._pending: "queue.Queue[str]" = queue.Queue()
        self._pending_order: List[str] = []
        self._order_lock = threading.Lock()
        self._worker = threading.Thread(target=self._worker_loop, daemon=True)
        self._worker.start()
        self._cleanup_old_jobs()

    def create_job(self, config: dict, source_file, template_file) -> Job:
        job_id = uuid.uuid4().hex
        job_upload_dir = UPLOAD_DIR / job_id
        job_upload_dir.mkdir(parents=True, exist_ok=True)

        source_path = job_upload_dir / f"source_{source_file.filename}"
        template_path = job_upload_dir / f"template_{template_file.filename}"

        with open(source_path, "wb") as f:
            shutil.copyfileobj(source_file.file, f)
        with open(template_path, "wb") as f:
            shutil.copyfileobj(template_file.file, f)

        result_dir = RESULT_DIR / job_id
        result_dir.mkdir(parents=True, exist_ok=True)

        job = Job(
            id=job_id,
            config=config,
            source_path=source_path,
            template_path=template_path,
            result_dir=result_dir,
        )
        self.jobs[job_id] = job
        with self._order_lock:
            self._pending_order.append(job_id)
        self._pending.put(job_id)
        return job

    def get(self, job_id: str) -> Optional[Job]:
        return self.jobs.get(job_id)

    def list_jobs(self) -> List[dict]:
        """Job 히스토리(최신순). 서버 프로세스가 살아있는 동안의 메모리 내 기록만 대상이다."""
        jobs = sorted(self.jobs.values(), key=lambda j: j.created_at, reverse=True)
        return [
            {
                "job_id": j.id,
                "status": j.status,
                "error": j.error,
                "created_at": j.created_at.isoformat(),
                "result_ready": j.status == JobStatus.DONE and j.result_zip_path.exists(),
                "source_filename": j.source_path.name.removeprefix("source_"),
                "queue_position": self.get_queue_position(j.id),
            }
            for j in jobs
        ]

    def get_queue_position(self, job_id: str) -> Optional[int]:
        """대기열에서 몇 번째인지(1부터). 대기 중이 아니면 None."""
        with self._order_lock:
            try:
                return self._pending_order.index(job_id) + 1
            except ValueError:
                return None

    def list_result_files(self, job_id: str) -> Optional[List[str]]:
        job = self.jobs.get(job_id)
        if job is None:
            return None
        return sorted(p.name for p in job.result_dir.glob("*.xlsx"))

    def cancel(self, job_id: str) -> bool:
        """데스크톱 버전의 writeFlag('0')에 대응 — 이 Job에만 취소 신호를 보낸다."""
        job = self.jobs.get(job_id)
        if job is None:
            return False
        job.cancel_event.set()
        return True

    def count_excel_processes(self) -> int:
        """
        작업 시작 전 "열려 있는 Excel 감지" 경고에 사용. 데스크톱 버전 사용법 문서
        (UsageUi.ui)의 "작업 전 Excel을 모두 종료하라"는 주의사항을 안내만 하는 대신,
        실제로 떠 있는지 확인해서 사용자가 판단할 수 있게 한다. 종료는 하지 않는다
        (force_kill_excel과 달리 읽기 전용).
        """
        import psutil

        return sum(
            1
            for proc in psutil.process_iter(["name"])
            if proc.info.get("name", "").lower() == "excel.exe"
        )

    def force_kill_excel(self) -> int:
        """
        데스크톱 버전의 psutil 기반 강제종료(excel.exe kill)에 대응.
        작업이 단일 워커로 직렬 처리되므로, 이 시점에 떠 있는 Excel 프로세스는
        많아야 현재 실행 중인 Job 하나에 속한다.
        """
        import psutil

        killed = 0
        for proc in psutil.process_iter(["name"]):
            if proc.info.get("name", "").lower() == "excel.exe":
                try:
                    proc.kill()
                    killed += 1
                except Exception:
                    pass
        return killed

    def _worker_loop(self) -> None:
        while True:
            job_id = self._pending.get()
            job = self.jobs.get(job_id)
            if job is None:
                continue
            self._run_job(job)

    def _run_job(self, job: Job) -> None:
        with self._order_lock:
            if job.id in self._pending_order:
                self._pending_order.remove(job.id)
        job.status = JobStatus.RUNNING

        job_logger = logging.getLogger(f"splitexcel.job.{job.id}")
        job_logger.setLevel(logging.DEBUG)
        job_logger.propagate = False
        handler = QueueLogHandler(job.log_queue)
        handler.setFormatter(logging.Formatter("%(asctime)s - %(message)s"))
        job_logger.addHandler(handler)

        pythoncom.CoInitialize()
        processor: Optional[ExcelSplitProcessor] = None
        try:
            cfg = job.config
            options = cfg.get("options", {})

            processor = ExcelSplitProcessor(
                source_path=str(job.source_path),
                template_path=str(job.template_path),
                result_path=str(job.result_dir),
                reulst_file_nm=cfg["result_file_nm"],
                result_file_date=cfg["result_file_date"],
                option_zoom_level1=options.get("zoom_level1", 120),
                option_zoom_level2=options.get("zoom_level2", 100),
                option_hide_guidelineYN=options.get("hide_guideline_yn", "Y"),
                option_formula_removePathYN=options.get("remove_formula_path_yn", "Y"),
                cancel_event=job.cancel_event,
                logger=job_logger,
            )

            sheet_tasks = [
                (t["sheet_name"], t["filter_index"], t["copy_range"], t["paste_range"])
                for t in cfg["sheet_tasks"]
            ]
            split_list = cfg["split_list"]
            if cfg.get("test_mode"):
                split_list = split_list[0:1]

            processor.process_sheets(split_list, sheet_tasks)

            if job.cancel_event.is_set():
                job.status = JobStatus.CANCELLED
            else:
                job.status = JobStatus.DONE
                self._zip_results(job)

        except Exception as e:
            job_logger.error(f"작업 실패: {e}")
            job.status = JobStatus.CANCELLED if job.cancel_event.is_set() else JobStatus.ERROR
            job.error = str(e)
        finally:
            if processor is not None and processor.excel:
                processor.quit_excel(processor.excel)
            pythoncom.CoUninitialize()
            job_logger.removeHandler(handler)
            job.log_queue.put(LOG_END_SENTINEL)
            job.finished = True

    def _zip_results(self, job: Job) -> None:
        with zipfile.ZipFile(job.result_zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
            for file in sorted(job.result_dir.glob("*.xlsx")):
                zf.write(file, arcname=file.name)

    def _cleanup_old_jobs(self, max_age_days: int = JOB_RETENTION_DAYS) -> int:
        """
        서버 시작 시 1회 실행되는 저장소 정리. storage/uploads/<job_id>/,
        storage/results/<job_id>/, storage/results/<job_id>.zip 중 전부 오래된(최신 수정시각
        기준 max_age_days 초과) 것만 지운다. 서버가 막 시작된 시점이라 self.jobs가 비어있으므로
        "지금 실행 중인 Job"과 충돌할 위험이 없다.
        """
        cutoff = datetime.datetime.now() - datetime.timedelta(days=max_age_days)
        job_ids = set()
        if UPLOAD_DIR.exists():
            job_ids.update(p.name for p in UPLOAD_DIR.iterdir() if p.is_dir())
        if RESULT_DIR.exists():
            for p in RESULT_DIR.iterdir():
                job_ids.add(p.stem if p.suffix == ".zip" else p.name)

        removed = 0
        for job_id in job_ids:
            paths = [
                p
                for p in (UPLOAD_DIR / job_id, RESULT_DIR / job_id, RESULT_DIR / f"{job_id}.zip")
                if p.exists()
            ]
            if not paths:
                continue
            newest = max(datetime.datetime.fromtimestamp(p.stat().st_mtime) for p in paths)
            if newest >= cutoff:
                continue
            for p in paths:
                try:
                    shutil.rmtree(p) if p.is_dir() else p.unlink()
                    removed += 1
                except OSError:
                    pass
        return removed


job_manager = JobManager()
