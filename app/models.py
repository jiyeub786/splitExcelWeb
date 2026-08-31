from typing import List

from pydantic import BaseModel, Field


class SheetTask(BaseModel):
    """splitExcel 데스크톱 버전의 table01(시트 작업 정의) 한 행에 대응."""

    sheet_name: str = Field(..., description="원본/양식에 공통으로 존재하는 시트명")
    filter_index: int = Field(..., description="AutoFilter를 적용할 열 번호(1부터 시작)")
    copy_range: str = Field(..., description="원본 시트에서 복사할 범위, 예: A1:D500")
    paste_range: str = Field(..., description="양식 시트에 붙여넣을 시작 위치, 예: A2")


class OptionsModel(BaseModel):
    """table03(옵션)에 대응."""

    zoom_level1: int = 120
    zoom_level2: int = 100
    hide_guideline_yn: str = "Y"
    remove_formula_path_yn: str = "Y"


class JobConfig(BaseModel):
    result_file_nm: str
    result_file_date: str
    sheet_tasks: List[SheetTask]
    split_list: List[str] = Field(..., description="table02(분리 대상 목록)에 대응")
    options: OptionsModel = OptionsModel()
    test_mode: bool = Field(False, description="True면 split_list 중 첫 항목만 실행(데스크톱 버전의 '테스트 시작')")
