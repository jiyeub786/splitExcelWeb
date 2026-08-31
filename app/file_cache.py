"""
업로드 파일 임시 캐시(미리보기 전용).

범위 선택 뷰어를 열 때마다, 분리 대상을 자동 추출할 때마다 브라우저가 원본/양식 파일 전체를
서버로 매번 재전송하면 네트워크 왕복이 낭비된다. 파일을 한 번 업로드하면 토큰을 내려주고,
이후 미리보기 계열 요청은 파일을 다시 첨부하지 않고 이 토큰만 보내면 재사용할 수 있다.

Job 실행에 실제로 쓰이는 원본/양식 파일 저장(job_manager.UPLOAD_DIR, 디스크에 영구 보관)과는
완전히 별개다 — 여기는 메모리에 잠깐만 들고 있다가 만료/청소되는 순수 미리보기용 캐시다.
"""

import threading
import time
import uuid
from typing import Dict, Optional, Tuple

_MAX_ENTRIES = 8
_TTL_SECONDS = 30 * 60

_lock = threading.Lock()
_store: Dict[str, Tuple[bytes, float]] = {}


def put(data: bytes) -> str:
    token = uuid.uuid4().hex
    with _lock:
        _evict_locked()
        _store[token] = (data, time.time())
    return token


def get(token: str) -> Optional[bytes]:
    with _lock:
        entry = _store.get(token)
        if entry is None:
            return None
        data, ts = entry
        if time.time() - ts > _TTL_SECONDS:
            del _store[token]
            return None
        return data


def _evict_locked() -> None:
    now = time.time()
    expired = [t for t, (_, ts) in _store.items() if now - ts > _TTL_SECONDS]
    for t in expired:
        del _store[t]
    while len(_store) >= _MAX_ENTRIES:
        oldest_token = min(_store.items(), key=lambda kv: kv[1][1])[0]
        del _store[oldest_token]
