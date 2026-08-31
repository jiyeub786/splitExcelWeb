"""
엑셀 미리보기용 유틸(python-calamine 기반).

Excel COM(win32com)을 띄우지 않고 파일 바이트만으로 즉시 응답할 수 있어야 하므로,
job_manager의 직렬 워커/Excel 프로세스와 무관하게 동시에 여러 요청을 처리해도 안전하다.
시트명 자동완성, 범위 선택 뷰어, 분리 대상 목록 자동 추출에 사용된다.

openpyxl 대신 python-calamine(Rust 기반 파서)을 사용한다: 실사용 대상 파일(수만~수십만 행,
수십MB)에서 openpyxl은 워크북을 여는 것만으로도 5초 이상 걸리는 반면, calamine은 같은 파일을
1초 내외로 열고 전체 시트를 읽어도 3초 안팎이라 UI 응답성 차이가 크다.
"""

import datetime
import re
from typing import BinaryIO, List

from python_calamine import CalamineWorkbook, WorksheetNotFound

_RANGE_RE = re.compile(r"^[A-Za-z]{1,3}[0-9]+(:[A-Za-z]{1,3}[0-9]+)?$")


def _format_cell(v) -> str:
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    if isinstance(v, (datetime.datetime, datetime.date, datetime.time)):
        return v.isoformat()
    return str(v)


def _column_index_from_letters(letters: str) -> int:
    idx = 0
    for ch in letters:
        if not ch.isalpha():
            raise ValueError(f"열 지정이 올바르지 않습니다: {letters}")
        idx = idx * 26 + (ord(ch.upper()) - ord("A") + 1)
    return idx


def _open_sheet(stream: BinaryIO, sheet_name: str):
    wb = CalamineWorkbook.from_filelike(stream)
    try:
        return wb.get_sheet_by_name(sheet_name)
    except WorksheetNotFound:
        raise ValueError(f"시트를 찾을 수 없습니다: {sheet_name}")


def list_sheet_names(stream: BinaryIO) -> List[str]:
    wb = CalamineWorkbook.from_filelike(stream)
    return list(wb.sheet_names)


def read_grid(stream: BinaryIO, sheet_name: str, max_rows: int = 200, max_cols: int = 40) -> dict:
    sheet = _open_sheet(stream, sheet_name)
    max_row = sheet.height
    max_column = sheet.width

    raw_rows = sheet.to_python(skip_empty_area=False, nrows=max_rows)
    rows = [[_format_cell(v) for v in row[:max_cols]] for row in raw_rows]

    # 실제 사용 범위가 요청한 max_rows/max_cols보다 짧아도(예: 헤더만 있고 데이터가 없는
    # 양식파일) 붙여넣기 위치를 헤더 아래 빈 칸에 클릭으로 지정할 수 있도록, 항상
    # max_rows x max_cols 크기로 빈 셀을 채워서 반환한다.
    for row in rows:
        if len(row) < max_cols:
            row.extend([""] * (max_cols - len(row)))
    while len(rows) < max_rows:
        rows.append([""] * max_cols)

    return {
        "sheet_name": sheet_name,
        "rows": rows,
        "max_row": max_row,
        "max_column": max_column,
        "truncated_rows": max_row > max_rows,
        "truncated_cols": max_column > max_cols,
    }


def extract_column_values(
    stream: BinaryIO, sheet_name: str, column: str, header_rows: int = 1
) -> List[str]:
    sheet = _open_sheet(stream, sheet_name)

    column = column.strip()
    try:
        col_idx = _column_index_from_letters(column) if column.isalpha() else int(column)
    except ValueError:
        raise ValueError(f"열 지정이 올바르지 않습니다: {column}")
    if col_idx < 1:
        raise ValueError(f"열 지정이 올바르지 않습니다: {column}")
    zero_idx = col_idx - 1
    header_rows = max(0, header_rows)

    values: List[str] = []
    seen = set()
    for i, row in enumerate(sheet.iter_rows()):
        if i < header_rows:
            continue
        if zero_idx >= len(row):
            continue
        v = row[zero_idx]
        if v is None:
            continue
        s = _format_cell(v).strip()
        if not s or s in seen:
            continue
        seen.add(s)
        values.append(s)
    return values


def validate_split_list(
    stream: BinaryIO, sheet_name: str, column: str, header_rows: int, split_list: List[str]
) -> dict:
    """분리 대상 목록의 각 값이 실제로 필터 열에 존재하는지 확인한다(오타로 빈 결과 파일이
    생기는 것을 사전에 방지하기 위함)."""
    actual_values = set(extract_column_values(stream, sheet_name, column, header_rows=header_rows))
    missing = [v for v in split_list if v not in actual_values]
    return {"missing": missing, "matched_count": len(split_list) - len(missing)}


def validate_range_syntax(range_str: str) -> bool:
    return bool(_RANGE_RE.match(range_str.strip()))


def dry_run_validate(
    source_stream: BinaryIO, template_stream: BinaryIO, sheet_tasks: List[dict]
) -> List[str]:
    """작업 시작 전 사전 점검: 시트명이 원본/양식 양쪽에 모두 있는지, 범위 문법이 올바른지 확인해
    Excel COM을 띄우기 전에 흔한 실수를 미리 잡아낸다."""
    errors: List[str] = []
    source_sheets = set(list_sheet_names(source_stream))
    template_sheets = set(list_sheet_names(template_stream))

    for i, task in enumerate(sheet_tasks, start=1):
        sheet_name = str(task.get("sheet_name", "")).strip()
        copy_range = str(task.get("copy_range", "")).strip()
        paste_range = str(task.get("paste_range", "")).strip()

        if not sheet_name:
            errors.append(f"{i}번째 행: 시트명이 비어 있습니다.")
        else:
            if sheet_name not in source_sheets:
                errors.append(f"{i}번째 행: 원본파일에 '{sheet_name}' 시트가 없습니다.")
            if sheet_name not in template_sheets:
                errors.append(f"{i}번째 행: 양식파일에 '{sheet_name}' 시트가 없습니다.")
        if not validate_range_syntax(copy_range):
            errors.append(f"{i}번째 행: 복사 범위 '{copy_range}' 형식이 올바르지 않습니다(예: A1:D500).")
        if not validate_range_syntax(paste_range):
            errors.append(f"{i}번째 행: 붙여넣기 위치 '{paste_range}' 형식이 올바르지 않습니다(예: A2).")
    return errors
