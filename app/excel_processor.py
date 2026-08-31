"""
splitExcel/package/ExcelSplitProcessor.py 를 웹(Job 단위 실행) 환경에 맞게 이식한 버전.

핵심 엑셀 처리 로직(win32com으로 Excel을 구동해 필터 -> 복사 -> 붙여넣기 -> 저장하는 흐름)은
원본 데스크톱 버전과 동일하게 유지한다. 바뀐 부분은 딱 두 가지뿐이다.

1. 취소 신호: 원본은 전역 파일(flag.txt)을 폴링했지만, 웹에서는 여러 Job이 동시에 큐에 쌓일 수
   있으므로 Job마다 독립적인 threading.Event(cancel_event)를 사용한다.
2. 로깅: 원본은 모듈 전역 로거('MyLogger')에 직접 기록했지만, 웹에서는 Job별 로그를 구분해서
   SSE로 스트리밍해야 하므로 로거 인스턴스를 생성자에서 주입받는다.
"""

import gc
import threading
import logging
import datetime
from typing import Optional

import win32com.client


class ExcelSplitProcessor:
    def __init__(
        self,
        source_path,
        template_path,
        result_path,
        reulst_file_nm,
        result_file_date,
        option_zoom_level1=120,  # 개요시트 zoom_level
        option_zoom_level2=100,  # 나머지 시트 zoom_level
        option_hide_guidelineYN="Y",  # 눈금선제거옵션 Y켜기 N 끄기
        option_formula_removePathYN="Y",
        cancel_event: Optional[threading.Event] = None,
        logger: Optional[logging.Logger] = None,
    ):
        self.excel = ""
        self.source_path = source_path
        self.template_path = template_path
        self.result_path = result_path
        self.reulst_file_nm = reulst_file_nm
        self.result_file_date = result_file_date
        self.option_zoom_level1 = int(option_zoom_level1)
        self.option_zoom_level2 = int(option_zoom_level2)
        self.option_hide_guidelineYN = option_hide_guidelineYN
        self.option_formula_removePathYN = option_formula_removePathYN
        self.target_workbook = ""
        self.source_workbook = ""

        self.cancel_event = cancel_event or threading.Event()
        self.logger = logger or logging.getLogger(__name__)

    def optimize_excel(self):
        """엑셀 성능 최적화를 위한 설정"""
        try:
            if self.excel.DisplayAlerts != False:
                self.excel.DisplayAlerts = False
            if self.excel.Calculation != -4135:
                self.excel.Calculation = -4135  # xlCalculationManual
            if self.excel.EnableEvents != False:
                self.excel.EnableEvents = False
            if self.excel.DisplayStatusBar != False:
                self.excel.DisplayStatusBar = False
            if self.excel.AskToUpdateLinks != False:
                self.excel.AskToUpdateLinks = False
        except Exception as e:
            self.logger.info(f"optimize_excel(): {e}")

    def restore_excel(self):
        """엑셀 성능 설정을 원래대로 복구"""
        try:
            if self.excel.Calculation != -4105:
                self.excel.Calculation = -4105  # xlCalculationAutomatic
        except Exception as e:
            self.logger.info(f"restore_excel(): {e}")

    def open_workbook(self, file_path):
        return self.excel.Workbooks.Open(file_path)

    def close_workbook(self, workbook, save_changes=False):
        try:
            if workbook is not None:
                workbook.Close(SaveChanges=save_changes)
        except Exception as e:
            self.logger.info(f"close_workbook(): {e}")

    def quit_excel(self, app):
        try:
            if app is not None:
                app.Quit()
        except Exception as e:
            self.logger.warning(f"quit_excel() error: {e}")
        finally:
            del app
            gc.collect()

    def filter_column(self, worksheet, filter_index, filter_value):
        try:
            worksheet.Range("A1").AutoFilter(Field=filter_index, Criteria1=filter_value)
        except Exception as e:
            self.logger.info(f"filter_column(): {e}")

    def clear_autofilter(self, worksheet):
        try:
            if worksheet.AutoFilterMode:
                worksheet.AutoFilterMode = False
        except Exception as e:
            self.logger.info(f"clear_autofilter(): {e}")

    def delete_rows_after_last_data(self, worksheet, column="A"):
        try:
            last_row = worksheet.Cells(worksheet.Rows.Count, column).End(-4162).Row  # xlUp
            total_rows = worksheet.Rows.Count
            self.clear_autofilter(worksheet)
            if last_row < total_rows:
                worksheet.Rows(f"{last_row + 1}:{total_rows}").Delete()
        except Exception as e:
            self.logger.info(f"delete_rows_after_last_data(): {e}")

    def remove_pathText(self, txt):
        import os

        fileName = self.source_path
        path = os.path.dirname(fileName).replace("/", "\\") + "\\"
        filename = "[" + os.path.basename(fileName).replace("/", "\\") + "]"
        return txt.replace(path, "").replace(filename, "").replace("@", "")

    def remove_external_references_from_worksheet(self, worksheet):
        try:
            used_range = worksheet.UsedRange
            formulas = used_range.Formula  # 2D array (row x col)

            row_count = used_range.Rows.Count
            col_count = used_range.Columns.Count

            updated_formulas = [[None for _ in range(col_count)] for _ in range(row_count)]
            changed = False

            for i in range(row_count):
                for j in range(col_count):
                    cell_formula = formulas[i][j]
                    if (
                        isinstance(cell_formula, str)
                        and "=" in cell_formula
                        and "[" in cell_formula
                        and "]" in cell_formula
                    ):
                        new_formula = self.remove_pathText(cell_formula)
                        updated_formulas[i][j] = new_formula
                        changed = True
                    else:
                        updated_formulas[i][j] = cell_formula

            if changed:
                used_range.Formula = updated_formulas

        except Exception as e:
            self.logger.info(f"remove_external_references_from_worksheet(): {e}")

    def copyAndPaste_sheet(self, soure_worksheet, target_worksheet, filter_index, filter_value, copy_range, paste_range):
        try:
            self.filter_column(soure_worksheet, filter_index, filter_value)
            filter_range = soure_worksheet.Range(copy_range)
            filter_range.Copy()
            target_worksheet.Range(paste_range).PasteSpecial()
            return 1
        except Exception as e:
            self.logger.info(f"copyAndPaste_sheet(): {e}")
            return 0

    def hide_gridlines(self):
        try:
            self.excel.ActiveWindow.DisplayGridlines = False
        except Exception as e:
            self.logger.info(f"hide_gridlines(): {e}")

    def set_init_workbook(self, workbook):
        try:
            sheet_cnt = workbook.Sheets.Count
            for sheet_num in range(sheet_cnt):
                sheet = workbook.Sheets(sheet_num + 1)

                if not sheet.Visible:
                    continue

                if self.option_hide_guidelineYN == "N":
                    self.hide_gridlines()

                sheet.Activate()
                sheet.Range("A1").Select()
                self.excel.ActiveWindow.ScrollRow = 1
                self.excel.ActiveWindow.ScrollColumn = 1

                if sheet_num + 1 == 1:
                    if self.option_zoom_level1 != 0:
                        self.excel.ActiveWindow.Zoom = self.option_zoom_level1
                else:
                    if self.option_zoom_level2 != 0:
                        self.excel.ActiveWindow.Zoom = self.option_zoom_level2

            for sheet_num in range(sheet_cnt):
                if workbook.Sheets(sheet_num + 1).Visible:
                    workbook.Sheets(sheet_num + 1).Activate()
                    break

        except Exception as e:
            self.logger.info(f"set_init_workbook(): {e}")

    def process_sheets(self, tgt_list, sheet_tasks):
        self.target_workbook = None
        self.soure_workbook = None

        try:
            self.logger.info("엑셀 분리작업을 시작합니다")

            if self.excel is None or not hasattr(self.excel, "Workbooks"):
                self.excel = win32com.client.Dispatch("Excel.Application")

            if self.soure_workbook is None or not hasattr(self.soure_workbook, "Close"):
                self.soure_workbook = self.open_workbook(self.source_path)

            self.optimize_excel()
            total_start_time = datetime.datetime.now()

            for i, tgt in enumerate(tgt_list):
                self.optimize_excel()
                each_start_time = datetime.datetime.now()
                self.target_workbook = self.open_workbook(self.template_path)

                for task in sheet_tasks:
                    soure_worksheet1 = self.soure_workbook.Sheets(task[0])
                    target_worksheet1 = self.target_workbook.Sheets(task[0])

                    if not self.cancel_event.is_set():
                        if 1 == self.copyAndPaste_sheet(
                            soure_worksheet1, target_worksheet1, task[1], tgt, task[2], task[3]
                        ):
                            self.delete_rows_after_last_data(target_worksheet1)
                            if self.option_formula_removePathYN == "Y":
                                self.remove_external_references_from_worksheet(target_worksheet1)

                    if self.cancel_event.is_set():
                        raise RuntimeError("종료 플래그 감지(취소 요청)")

                self.set_init_workbook(self.target_workbook)
                self.restore_excel()

                reulst_file_nm1 = f"{self.reulst_file_nm}_{str(i + 1).zfill(2)}_{tgt}_{self.result_file_date}.xlsx"
                self.target_workbook.SaveAs(f"{self.result_path}/{reulst_file_nm1}")
                self.close_workbook(self.target_workbook, save_changes=False)
                self.target_workbook = None

                elapsed_time = datetime.datetime.now() - each_start_time
                self.logger.info(f"{tgt} 완료. 처리 시간: {int(elapsed_time.total_seconds())}초")

            self.close_workbook(self.soure_workbook, save_changes=False)
            self.soure_workbook = None

            elapsed_time = datetime.datetime.now() - total_start_time
            self.logger.info(f"모든작업 완료. 총 처리 시간: {int(elapsed_time.total_seconds())}초")

        except Exception as e:
            self.logger.error(f"process_sheets() Error processing: {e}")
            self.target_workbook = None
            self.close_workbook(self.soure_workbook, save_changes=False)
            self.soure_workbook = None
            raise
