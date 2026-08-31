# CLAUDE.md

이 파일은 이 저장소에서 작업할 때 Claude Code(claude.ai/code)에게 지침을 제공합니다.

## 저장소 개요

원본 엑셀 워크북을 대상값별로 여러 출력 워크북으로 분리하는 FastAPI 웹 애플리케이션
("SplitExcel Web")입니다. PySide6 데스크톱 앱이었던 원본 "SplitExcel"을 웹으로 이식한 버전으로,
Excel COM 자동화 로직 자체는 그대로 유지하면서 UI와 다중 작업(Job) 실행/취소/로깅 방식만 웹에
맞게 재설계했습니다. Windows + Microsoft Excel 설치 환경에서만 동작합니다.

- `docs/design.md` — 프런트엔드가 따르는 디자인 시스템 분석 문서(Binance 스타일 분석: 색상/
  타이포그래피/컴포넌트 토큰). `static/` 아래 CSS/HTML을 수정할 때 참고하세요.
- `docs/design2.md` — `design.md`의 라이트 테마 변형 토큰 문서. 헤더의 🌙/☀️ 토글 버튼이 전환하는
  라이트 모드 색상은 이 문서 기준입니다. 브랜드 액센트/타이포그래피/반경은 `design.md`와 공유하고
  캔버스·표면·텍스트 톤만 다시 정의합니다.
- `docs/plan.md` — 초기 개선 아이디어 요청 메모.
- `개선아이디어.md` — 위 요청에 대한 개선 아이디어 정리 문서(버전관리 대상 아님, `.gitignore` 참고).

이 저장소에는 자동화된 테스트나 린터/포매터 설정이 없습니다.

## 명령어

> ⚠️ **`python` 명령어를 그냥 쓰지 마세요.** PC에 Python이 여러 버전 설치되어 있을 수 있습니다.
> `fastapi`, `pydantic` 등 최신 패키지는 구버전 Python을 지원하지 않으므로 반드시 Python 3.13
> 이상으로 venv를 만드세요(Windows라면 **`py -3.13`**으로 버전을 지정). `py -0p`로 이 PC에 설치된
> Python 버전과 경로를 확인할 수 있습니다.

```
py -3.13 -m venv .venv
.venv\Scripts\activate
python -m pip install --upgrade pip
pip install -r requirements.txt

# pywin32 설치 후 1회 필요 (win32com 관련 DLL 등록)
python .venv\Scripts\pywin32_postinstall.py -install

# 개발 서버 실행 (http://127.0.0.1:8000)
uvicorn app.main:app --reload --port 8000
```

이미 구버전 Python으로 `.venv`를 만들어버렸다면 그 폴더를 삭제하고 위 명령으로 다시 만들면 됩니다
(`.venv`는 언제든 재생성 가능한 산출물이라 지워도 안전합니다).

## 아키텍처

**진입점 / API** (`app/main.py`): FastAPI 앱이 REST 엔드포인트와 정적 프런트엔드(`static/`)를 함께
서빙합니다.
- `POST /api/jobs` — 원본파일 + 양식파일 업로드와 `JobConfig`(JSON) 설정을 받아 `job_manager`에
  Job을 등록.
- `GET /api/jobs/{job_id}` — Job 상태 조회.
- `POST /api/jobs/{job_id}/cancel` — 협조적 취소.
- `POST /api/jobs/{job_id}/force-stop-excel` — `excel.exe` 프로세스 강제 종료.
- `GET /api/jobs/{job_id}/logs` — Server-Sent Events로 실행 로그 실시간 스트리밍.
- `GET /api/jobs/{job_id}/result` — 완료된 결과 zip 다운로드.

**Job 생명주기** (`app/job_manager.py`): Excel COM은 동시 다중 자동화에 취약하므로, 등록된 Job은
단일 백그라운드 워커 스레드(`JobManager._worker_loop`)가 큐에서 하나씩 꺼내 **직렬로** 처리합니다.
각 Job은 자체 `threading.Event`(취소 신호)와 자체 `queue.Queue`(로그 라인, `QueueLogHandler`를
통해 `logging` 레코드를 적재 → SSE로 스트리밍)를 가집니다. `_run_job`은 Job 설정으로
`ExcelSplitProcessor`를 생성해 `process_sheets`를 호출하고, 완료되면
`storage/results/<job_id>/*.xlsx`를 `<job_id>.zip`으로 압축합니다. `force_kill_excel`은 `psutil`로
`excel.exe` 프로세스를 찾아 kill합니다.

**핵심 로직** (`app/excel_processor.py`): `win32com.client`로 실제 Excel 인스턴스를 COM으로
구동합니다(파일 포맷 라이브러리가 아님). 분리 목록의 각 대상값마다 템플릿 워크북의 새 사본을 열고,
각 시트 작업(`sheet_name`, `filter_index`, `copy_range`, `paste_range`)마다 원본 시트를 필터링하고
필터링된 범위를 템플릿에 복사/붙여넣기한 뒤, 끝부분의 빈 행을 잘라내고
(`delete_rows_after_last_data`), 선택적으로 수식에서 외부 파일 경로 참조를 제거
(`remove_external_references_from_worksheet`)한 후 결과를
`{result_file_nm}_{seq}_{target}_{result_file_date}.xlsx`로 저장합니다. `optimize_excel`/
`restore_excel`이 배치 작업 전후로 Excel 앱 레벨 설정(경고창, 계산 모드, 이벤트, 상태 표시줄)을
토글해 COM 자동화 속도를 높입니다. 실제 Excel 프로세스를 구동하므로 Excel이 설치된 Windows에서만
동작합니다.

**요청 스키마** (`app/models.py`, Pydantic): `SheetTask`(시트 작업 한 행에 대응),
`OptionsModel`(줌 레벨 2종, 눈금선 제거 여부, 수식 경로 참조 제거 여부), `JobConfig`
(`result_file_nm`, `result_file_date`, `sheet_tasks`, `split_list`(분리 대상 목록), `options`,
`test_mode`).

**프런트엔드** (`static/index.html`, `style.css`, `app.js`): 별도 빌드 단계가 없는 순수 HTML/CSS/
바닐라 JS. `app.js`가 시트 작업 테이블 행 추가/삭제, 분리 대상 목록/옵션 수집, 설정 JSON
다운로드/업로드, `/api/jobs` 제출, `EventSource`를 통한 로그 스트림 구독, 폴링을 통한 상태 갱신,
취소/강제종료 버튼을 담당합니다. 시각 스타일(`style.css`)은 `docs/design.md` 디자인 시스템
(다크 캔버스 + 단일 노란색 액센트 + BinanceNova/BinancePlex 타이포그래피 조합)을 따릅니다 — 새
컴포넌트를 추가하거나 스타일을 바꿀 때는 `docs/design.md`의 토큰(`{colors.*}`, `{typography.*}`,
`{rounded.*}`, `{spacing.*}`)을 참고하세요. 헤더의 🌙/☀️ 버튼으로 전환하는 라이트 테마 색상은
`docs/design2.md`를 따르며, `style.css`는 `:root`(다크 기본값)와 `:root[data-theme="light"]`
(라이트 오버라이드)로 토큰을 나눠 정의합니다.

**저장소 레이아웃**: 업로드 원본/결과 파일은 `storage/uploads/<job_id>/`, `storage/results/<job_id>/`
+ `<job_id>.zip`에 런타임에 생성됩니다(둘 다 버전관리 대상 아님, 정리 로직 없음 — 오래된 Job
디렉터리는 수동으로 지워야 함).

## 알려진 제약

- **동시성**: Excel COM 특성상 여러 요청을 동시에 안전하게 처리하기 어려워, 현재는 워커 스레드
  1개가 Job을 순차 처리합니다. 동시 사용자가 늘어나면 대기 시간이 길어질 수 있습니다.
- **세션 환경**: 서버를 Windows 서비스로 돌릴 경우 인터랙티브 세션이 아니어서 Excel이 정상적으로
  뜨지 않을 수 있습니다. 일반 사용자 세션에서 `uvicorn`을 직접 실행하는 것을 전제로 합니다.
- **인증 없음**: 사내 전용 도구를 가정한 1차 이식이라 별도 인증/권한 처리는 포함하지 않았습니다.

## 작업 원칙 (행동 가이드라인)

이 저장소에 적용되는 일반 행동 지침입니다. 출처:
[andrej-karpathy-skills/CLAUDE.md](https://github.com/multica-ai/andrej-karpathy-skills/blob/main/CLAUDE.md).
아래 원칙은 속도보다 신중함 쪽으로 치우쳐 있습니다 — 사소한 작업에는 판단해서 유연하게 적용하세요.

### 1. 코딩 전에 먼저 생각하기

추측하지 말고, 헷갈리는 부분을 숨기지 말고, 트레이드오프를 드러낼 것.

- 구현 전에 가정을 명시적으로 밝히세요. 확신이 없으면 물어보세요.
- 여러 해석이 가능하면 조용히 하나를 고르지 말고 제시하세요.
- 더 단순한 접근이 있다면 말하세요. 필요하면 반박하세요.
- 불명확한 부분이 있으면 멈추고, 무엇이 헷갈리는지 짚고, 질문하세요.

### 2. 단순함 우선

문제를 해결하는 최소한의 코드만. 추측성 코드는 넣지 않는다.

- 요청하지 않은 기능은 추가하지 않습니다.
- 한 번만 쓰는 코드에 추상화를 만들지 않습니다.
- 요청받지 않은 "유연성"/"설정 가능성"을 넣지 않습니다.
- 일어날 수 없는 상황에 대한 에러 처리를 넣지 않습니다.
- 200줄을 썼는데 50줄로 될 수 있다면 다시 쓰세요.

"시니어 엔지니어가 보면 과하게 복잡하다고 할까?" 스스로 물어보고, 그렇다면 단순화하세요.

### 3. 외과적 수정 (Surgical Changes)

반드시 건드려야 할 것만 건드리고, 내가 어질러 놓은 것만 치운다.

- 기존 코드를 수정할 때 인접한 코드/주석/포맷팅을 "개선"하지 않습니다.
- 망가지지 않은 것을 리팩터링하지 않습니다.
- 내 취향과 다르더라도 기존 스타일을 따릅니다.
- 관련 없는 죽은 코드를 발견하면 언급만 하고 삭제하지 않습니다.
- 내 변경으로 인해 안 쓰게 된 import/변수/함수는 정리합니다. 원래 있던 죽은 코드는 요청 없이
  지우지 않습니다.
- 기준: 변경된 모든 줄이 사용자의 요청과 직접 연결되어야 합니다.

### 4. 목표 지향 실행

성공 기준을 정의하고, 검증될 때까지 반복한다.

- 작업을 검증 가능한 목표로 바꾸세요. 예: "검증 로직 추가" → "잘못된 입력에 대한 테스트를 작성한
  뒤 통과시키기", "버그 수정" → "버그를 재현하는 테스트를 작성한 뒤 통과시키기", "X 리팩터링" →
  "리팩터링 전후로 테스트가 통과하는지 확인".
- 여러 단계로 이루어진 작업은 간단한 계획을 먼저 제시하세요 (단계 → 검증 방법).
- 이 저장소에는 자동화된 테스트가 없으므로, "검증"은 수동 실행/재현 절차로 대체될 수 있습니다.

## 작업 방식

응답이 (분석, 계획, 요약, 가이드, 생성된 코드 등) 한 줄짜리 답변이 아닌 실질적인 결과물을 만들어낼
때는, 채팅에만 남기지 말고 파일로 작성한 뒤 `SendUserFile`로 전달하세요. 결과물이 코드베이스에
속한다면 이 프로젝트 폴더에, 세션에 한정된 메모라면 스크래치패드에 저장하세요. 채팅 응답 자체는
짧게 유지하고 세부 내용은 파일을 가리키도록 하세요. 사소하고 짧은 답변에는 이 방식이 필요하지
않습니다.
