from __future__ import annotations

import csv
import gc
import math
import os
import re
import time
from dataclasses import dataclass
from datetime import date, datetime, timedelta
from decimal import Decimal
from pathlib import Path
from typing import Any, Callable, Iterable, Sequence

from Modules.Shipments.DailyInbound import (
    extract_dispatch_numbers,
    normalize_dispatch_number,
)


LogCallback = Callable[[str], None]


class ExcelImportError(RuntimeError):
    """Raised when downloaded Milkrun values cannot be written safely."""


class ExcelWorkbookOpenError(ExcelImportError):
    """Raised when UnHelper requires the linked workbook to be closed."""


class ExcelImportCancelled(RuntimeError):
    """Raised before the target workbook mutation starts."""


@dataclass(frozen=True)
class MilkrunExcelImportResult:
    source_file: Path
    target_workbook: Path
    sheet_name: str
    rows: int
    columns: int
    dispatch_numbers: tuple[str, ...]
    filtered_rows: int = 0


class MilkrunExcelImporter:
    """Replace the configured Milkrun RAW area with downloaded values.

    ``Value2`` assignment intentionally mirrors Excel's paste-values operation.
    It does not transfer formulas, formatting, column widths, or clipboard state.
    """

    TARGET_SHEET = "Raw_밀크런"
    CLEAR_RANGE = "C1:P1000"
    START_ROW = 1
    START_COLUMN = 3
    MAX_ROWS = 1000
    MAX_SOURCE_ROWS = 10_000
    MAX_COLUMNS = 14
    MAX_SOURCE_BYTES = 50 * 1024 * 1024
    TARGET_FILENAME_MARKER = "입고스케줄관리"
    TARGET_EXTENSIONS = frozenset({".xlsx", ".xlsm", ".xlsb", ".xls"})
    SOURCE_EXCEL_EXTENSIONS = TARGET_EXTENSIONS
    SOURCE_TEXT_EXTENSIONS = frozenset({".csv", ".txt", ".tsv"})
    _COM_OPERATION_MAX_RETRIES = 5
    _COM_RETRY_BASE_DELAY_SECONDS = 0.5
    _COM_READY_MAX_CHECKS = 4
    _COM_READY_BASE_DELAY_SECONDS = 0.2
    _COM_BUSY_HRESULTS = frozenset({
        -2147418111,  # RPC_E_CALL_REJECTED
        -2147417846,  # RPC_E_SERVERCALL_RETRYLATER
    })
    _COM_BUSY_RETRY_MARKERS = (
        "-2147418111",
        "-2147417846",
        "0x80010001",
        "0x8001010a",
        "RPC_E_CALL_REJECTED",
        "RPC_E_SERVERCALL_RETRYLATER",
        "피호출자가 호출을 거부",
        "Call was rejected by callee",
        "The message filter indicated that the application is busy",
        "메시지 필터",
    )
    _SAVE_RETRY_MARKERS = (
        "-2146827284",
        "문서가 저장되지",
        "Document not saved",
    )

    def __init__(
        self,
        log: LogCallback | None = None,
        *,
        com_client: Any | None = None,
        pythoncom_module: Any | None = None,
        reject_open_target: bool = False,
    ):
        self.log = log or (lambda _message: None)
        self._com_client = com_client
        self._pythoncom = pythoncom_module
        self._reject_open_target = reject_open_target

    @classmethod
    def validate_target_path(cls, value: str | Path) -> Path:
        path = Path(value).expanduser()
        if not path.is_file():
            raise ExcelImportError(f"연결된 Excel 파일을 찾을 수 없습니다.\n{path}")
        if path.suffix.lower() not in cls.TARGET_EXTENSIONS:
            extensions = ", ".join(sorted(cls.TARGET_EXTENSIONS))
            raise ExcelImportError(f"지원하지 않는 Excel 파일 형식입니다: {path.suffix or '(확장자 없음)'}\n지원 형식: {extensions}")
        if cls.TARGET_FILENAME_MARKER not in path.stem:
            raise ExcelImportError(
                f"연결할 Excel 파일 이름에 '{cls.TARGET_FILENAME_MARKER}'가 포함되어 있어야 합니다.\n"
                f"선택한 파일: {path.name}"
            )
        return path.resolve()

    @classmethod
    def validate_source_path(cls, value: str | Path) -> Path:
        path = Path(value).expanduser()
        if not path.is_file():
            raise ExcelImportError(f"다운로드 파일을 찾을 수 없습니다.\n{path}")
        supported = cls.SOURCE_EXCEL_EXTENSIONS | cls.SOURCE_TEXT_EXTENSIONS
        if path.suffix.lower() not in supported:
            extensions = ", ".join(sorted(supported))
            raise ExcelImportError(
                f"다운로드 파일 형식을 읽을 수 없습니다: {path.suffix or '(확장자 없음)'}\n"
                f"지원 형식: {extensions}"
            )
        if path.stat().st_size > cls.MAX_SOURCE_BYTES:
            raise ExcelImportError(
                "다운로드 파일이 안전하게 처리할 수 있는 크기를 초과했습니다.\n"
                f"최대 크기: {cls.MAX_SOURCE_BYTES // (1024 * 1024)}MB"
            )
        return path.resolve()

    def validate_workbook(self, target_workbook: str | Path) -> Path:
        """Validate the configured workbook in the worker before Chrome starts."""
        target_path = self.validate_target_path(target_workbook)
        pythoncom, com_client = self._load_com_modules()
        coinitialized = False
        excel = None
        target = None
        owns_excel_instance = False
        owns_target_workbook = False
        try:
            pythoncom.CoInitialize()
            coinitialized = True
            excel, target, owns_excel_instance, owns_target_workbook = self._open_target_workbook(
                com_client,
                target_path,
            )
            self._ensure_target_ready(excel, target)
            self.log(f"연결된 Excel 파일과 '{self.TARGET_SHEET}' 시트를 확인했습니다.")
            return target_path
        finally:
            if target is not None and owns_target_workbook:
                try:
                    target.Close(SaveChanges=False)
                except Exception:
                    pass
            if excel is not None and owns_excel_instance:
                try:
                    excel.Quit()
                except Exception:
                    pass
            target = None
            excel = None
            gc.collect()
            if coinitialized:
                try:
                    pythoncom.CoUninitialize()
                except Exception:
                    pass

    def import_values(
        self,
        source_file: str | Path,
        target_workbook: str | Path,
        *,
        cancel_requested: Callable[[], bool] | None = None,
        exclude_arrival_date: date | None = None,
    ) -> MilkrunExcelImportResult:
        source_path = self.validate_source_path(source_file)
        target_path = self.validate_target_path(target_workbook)
        if self._same_path(source_path, target_path):
            raise ExcelImportError("다운로드 파일과 연결된 Excel 파일이 같습니다. 서로 다른 파일을 선택해 주세요.")
        source_kind = self._classify_source(source_path)

        pythoncom, com_client = self._load_com_modules()
        coinitialized = False
        excel = None
        target = None
        source_book = None
        owns_excel_instance = False
        owns_target_workbook = False

        try:
            pythoncom.CoInitialize()
            coinitialized = True

            excel, target, owns_excel_instance, owns_target_workbook = self._open_target_workbook(
                com_client,
                target_path,
            )
            sheet = self._ensure_target_ready(
                excel,
                target,
                cancel_requested=cancel_requested,
            )
            self.log(f"다운로드 파일의 값 데이터를 읽습니다: {source_path.name}")
            if source_kind == "text":
                values = self._read_delimited_values(source_path)
            else:
                source_book = self._open_source_workbook(excel, source_path)
                values = self._read_excel_values(source_book)

            source_rows = len(values)
            source_columns = max((len(row) for row in values), default=0)
            if source_rows == 0 or source_columns == 0:
                raise ExcelImportError("다운로드 파일에 붙여넣을 값이 없습니다.")
            if source_rows > self.MAX_SOURCE_ROWS:
                raise ExcelImportError(
                    f"다운로드 데이터가 안전 처리 한도 {self.MAX_SOURCE_ROWS:,}행을 초과합니다. "
                    "기존 값을 지우지 않았습니다."
                )
            if source_columns > self.MAX_COLUMNS:
                raise ExcelImportError(
                    "다운로드 데이터가 대상 범위보다 큽니다. 기존 값을 지우지 않았습니다.\n"
                    f"다운로드: {source_rows}행 × {source_columns}열 / "
                    f"대상: {self.MAX_ROWS}행 × {self.MAX_COLUMNS}열"
                )
            filtered_rows = 0
            if exclude_arrival_date is not None:
                values, filtered_rows = self._filter_arrival_date_rows(
                    values,
                    exclude_arrival_date,
                )
                self.log(f"제외 대상 입고일 데이터 {filtered_rows}행을 붙여넣기에서 제외했습니다.")

            rows = len(values)
            columns = max((len(row) for row in values), default=0)
            if rows > self.MAX_ROWS:
                size_description = (
                    "입고일이 기준일 전날인 행을 제외한 데이터"
                    if exclude_arrival_date is not None
                    else "반영할 다운로드 데이터"
                )
                raise ExcelImportError(
                    f"{size_description}가 대상 범위보다 큽니다. "
                    "기존 값을 지우지 않았습니다.\n"
                    f"반영 대상: {rows}행 × {columns}열 / "
                    f"대상: {self.MAX_ROWS}행 × {self.MAX_COLUMNS}열"
                )
            if cancel_requested is not None and cancel_requested():
                raise ExcelImportCancelled("사용자가 작업을 중지했습니다.")

            normalized_values = self._rectangularize(values, columns)
            dispatch_numbers, import_metadata = self._extract_import_metadata(
                normalized_values,
                exclude_arrival_date=exclude_arrival_date,
            )
            clear_range = self._run_excel_com_operation(
                excel,
                lambda: sheet.Range(self.CLEAR_RANGE),
                "대상 범위 확인",
                cancel_requested=cancel_requested,
            )
            original_attribute = "Value2"
            original_contents = self._run_excel_com_operation(
                excel,
                lambda: clear_range.Value2,
                "기존 값 백업",
            )
            try:
                original_contents = self._run_excel_com_operation(
                    excel,
                    lambda: clear_range.Formula2,
                    "기존 수식 백업",
                )
                original_attribute = "Formula2"
            except Exception as formula2_error:
                if self._is_excel_busy_error(formula2_error):
                    raise
                try:
                    original_contents = self._run_excel_com_operation(
                        excel,
                        lambda: clear_range.Formula,
                        "기존 수식 백업",
                    )
                    original_attribute = "Formula"
                except Exception as formula_error:
                    if self._is_excel_busy_error(formula_error):
                        raise
                    pass
            changed_target = False
            try:
                self.log(f"{self.TARGET_SHEET}!{self.CLEAR_RANGE}의 기존 값만 지웁니다.")
                self._run_excel_com_operation(
                    excel,
                    clear_range.ClearContents,
                    "기존 값 지우기",
                    cancel_requested=cancel_requested,
                )
                changed_target = True
                destination = self._run_excel_com_operation(
                    excel,
                    lambda: sheet.Range(
                        sheet.Cells(self.START_ROW, self.START_COLUMN),
                        sheet.Cells(
                            self.START_ROW + rows - 1,
                            self.START_COLUMN + columns - 1,
                        ),
                    ),
                    "붙여넣을 범위 확인",
                    cancel_requested=cancel_requested,
                )
                self._run_excel_com_operation(
                    excel,
                    lambda: setattr(destination, "Value2", normalized_values),
                    "값 붙여넣기",
                    cancel_requested=cancel_requested,
                )
                if cancel_requested is not None and cancel_requested():
                    raise ExcelImportCancelled("사용자가 작업을 중지했습니다.")
                self.log(f"{self.TARGET_SHEET}!C1부터 {rows}행 × {columns}열을 값으로 붙여넣었습니다.")
                self._save_workbook(
                    excel,
                    target,
                    suppress_display_alerts=owns_excel_instance,
                    cancel_requested=cancel_requested,
                )
            except Exception as write_error:
                rollback_error = None
                if changed_target:
                    try:
                        self._run_excel_com_operation(
                            excel,
                            lambda: setattr(clear_range, original_attribute, original_contents),
                            "기존 값 복원",
                        )
                    except Exception as exc:
                        rollback_error = exc
                if rollback_error is not None:
                    raise ExcelImportError(
                        "Excel 반영 실패 후 기존 값 복원에도 실패했습니다. 대상 파일을 저장하지 말고 닫은 뒤 다시 열어 주세요.\n"
                        f"반영 오류: {write_error}\n복원 오류: {rollback_error}"
                    ) from write_error
                raise

            return self._build_import_result(
                source_file=source_path,
                target_workbook=target_path,
                sheet_name=self.TARGET_SHEET,
                rows=rows,
                columns=columns,
                dispatch_numbers=dispatch_numbers,
                filtered_rows=filtered_rows,
                import_metadata=import_metadata,
            )
        except (ExcelImportCancelled, ExcelImportError):
            raise
        except Exception as exc:
            raise ExcelImportError(f"Excel 값 붙여넣기에 실패했습니다.\n{exc}") from exc
        finally:
            if source_book is not None:
                try:
                    source_book.Close(SaveChanges=False)
                except Exception:
                    pass
            if target is not None and owns_target_workbook:
                try:
                    target.Close(SaveChanges=False)
                except Exception:
                    pass
            if excel is not None and owns_excel_instance:
                try:
                    excel.Quit()
                except Exception:
                    pass

            source_book = None
            target = None
            excel = None
            gc.collect()
            if coinitialized:
                try:
                    pythoncom.CoUninitialize()
                except Exception:
                    pass

    def _extract_import_metadata(
        self,
        values: tuple[tuple[Any, ...], ...],
        *,
        exclude_arrival_date: date | None,
    ) -> tuple[tuple[str, ...], Any]:
        """Validate Milkrun lookup keys and return optional result metadata.

        Subclasses can override this hook for another inbound-booking format
        while sharing the source checks, COM mutation, rollback, and save path.
        """

        dispatch_numbers = extract_dispatch_numbers(values)
        if exclude_arrival_date is not None and len(values) > 1 and not dispatch_numbers:
            raise ExcelImportError(
                "기준일 입고 데이터의 A열에서 Milkrun 배차번호를 찾지 못했습니다. "
                "기존 값을 지우지 않았습니다."
            )
        return dispatch_numbers, None

    def _build_import_result(
        self,
        *,
        source_file: Path,
        target_workbook: Path,
        sheet_name: str,
        rows: int,
        columns: int,
        dispatch_numbers: tuple[str, ...],
        filtered_rows: int,
        import_metadata: Any,
    ) -> MilkrunExcelImportResult:
        del import_metadata
        return MilkrunExcelImportResult(
            source_file=source_file,
            target_workbook=target_workbook,
            sheet_name=sheet_name,
            rows=rows,
            columns=columns,
            dispatch_numbers=dispatch_numbers,
            filtered_rows=filtered_rows,
        )

    def _load_com_modules(self) -> tuple[Any, Any]:
        if self._pythoncom is not None and self._com_client is not None:
            return self._pythoncom, self._com_client
        try:
            import pythoncom
            import win32com.client
        except ImportError as exc:
            raise ExcelImportError(
                "Excel 자동화 구성요소를 불러오지 못했습니다. UnHelper를 다시 설치해 주세요."
            ) from exc
        return pythoncom, win32com.client

    def _open_target_workbook(
        self,
        com_client: Any,
        target_path: Path,
    ) -> tuple[Any, Any, bool, bool]:
        active_excel = None
        try:
            active_excel = com_client.GetActiveObject("Excel.Application")
        except Exception:
            pass

        if active_excel is not None:
            open_target = self._find_open_workbook(active_excel, target_path)
            if open_target is not None:
                if self._reject_open_target:
                    raise ExcelWorkbookOpenError(
                        "연결된 입고스케줄 Excel 파일이 열려 있습니다.\n"
                        "Excel에서 해당 파일을 닫은 뒤 다시 시도해 주세요.\n"
                        f"파일: {target_path.name}"
                    )
                self.log(f"열려 있는 Excel 파일에 연결했습니다: {target_path.name}")
                return active_excel, open_target, False, False

        excel = None
        try:
            excel = com_client.DispatchEx("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False
            target = self._workbooks_open_with_macros_disabled(
                excel,
                str(target_path),
                UpdateLinks=0,
                ReadOnly=False,
                IgnoreReadOnlyRecommended=True,
                AddToMru=False,
            )
            self.log(f"연결된 Excel 파일을 백그라운드에서 열었습니다: {target_path.name}")
            return excel, target, True, True
        except Exception as exc:
            if excel is not None:
                try:
                    excel.Quit()
                except Exception:
                    pass
            raise ExcelImportError(
                "연결된 Excel 파일을 열지 못했습니다. 파일이 다른 프로그램에서 잠겨 있지 않은지 확인해 주세요.\n"
                f"{exc}"
            ) from exc

    @classmethod
    def _find_open_workbook(cls, excel: Any, target_path: Path) -> Any | None:
        for workbook in excel.Workbooks:
            try:
                if cls._same_path(Path(str(workbook.FullName)), target_path):
                    return workbook
            except Exception:
                continue
        return None

    def _get_target_sheet(
        self,
        excel: Any,
        workbook: Any,
        *,
        cancel_requested: Callable[[], bool] | None = None,
    ) -> Any:
        try:
            return self._run_excel_com_operation(
                excel,
                lambda: workbook.Worksheets(self.TARGET_SHEET),
                f"'{self.TARGET_SHEET}' 시트 확인",
                cancel_requested=cancel_requested,
            )
        except ExcelImportCancelled:
            raise
        except Exception as exc:
            if self._is_excel_busy_error(exc):
                raise ExcelImportError(
                    "Excel이 계속 사용 중이어 연결된 파일의 "
                    f"'{self.TARGET_SHEET}' 시트를 확인하지 못했습니다. "
                    "Excel의 편집 중인 셀이나 팝업창을 닫고 다시 시도해 주세요."
                ) from exc
            raise ExcelImportError(
                f"연결된 Excel 파일에 '{self.TARGET_SHEET}' 시트가 없습니다. 기존 값은 변경하지 않았습니다."
            ) from exc

    def _ensure_target_ready(
        self,
        excel: Any,
        workbook: Any,
        *,
        cancel_requested: Callable[[], bool] | None = None,
    ) -> Any:
        try:
            read_only = self._run_excel_com_operation(
                excel,
                lambda: bool(getattr(workbook, "ReadOnly", False)),
                "읽기 전용 상태 확인",
                cancel_requested=cancel_requested,
            )
        except ExcelImportCancelled:
            raise
        except Exception as exc:
            if self._is_excel_busy_error(exc):
                raise ExcelImportError(
                    "Excel이 계속 사용 중이어 연결된 파일의 편집 상태를 "
                    "확인하지 못했습니다. Excel의 편집 중인 셀이나 팝업창을 닫고 "
                    "다시 시도해 주세요."
                ) from exc
            raise
        if read_only:
            raise ExcelImportError(
                "연결된 Excel 파일이 읽기 전용입니다. Excel에서 '편집 사용'으로 연 뒤 다시 시도해 주세요."
            )
        try:
            saved = self._run_excel_com_operation(
                excel,
                lambda: bool(getattr(workbook, "Saved", True)),
                "저장 상태 확인",
                cancel_requested=cancel_requested,
            )
        except ExcelImportCancelled:
            raise
        except Exception as exc:
            if self._is_excel_busy_error(exc):
                raise ExcelImportError(
                    "Excel이 계속 사용 중이어 연결된 파일의 저장 상태를 "
                    "확인하지 못했습니다. Excel의 편집 중인 셀이나 팝업창을 닫고 "
                    "다시 시도해 주세요."
                ) from exc
            raise
        if not saved:
            raise ExcelImportError(
                "연결된 Excel 파일에 저장하지 않은 변경사항이 있습니다. Excel에서 먼저 저장한 뒤 다시 시도해 주세요."
            )
        return self._get_target_sheet(
            excel,
            workbook,
            cancel_requested=cancel_requested,
        )

    @classmethod
    def _open_source_workbook(cls, excel: Any, source_path: Path) -> Any:
        try:
            return cls._workbooks_open_with_macros_disabled(
                excel,
                str(source_path),
                UpdateLinks=0,
                ReadOnly=True,
                IgnoreReadOnlyRecommended=True,
                AddToMru=False,
            )
        except Exception as exc:
            raise ExcelImportError(f"다운로드 Excel 파일을 열지 못했습니다.\n{exc}") from exc

    @staticmethod
    def _workbooks_open_with_macros_disabled(excel: Any, path: str, **kwargs) -> Any:
        previous_security = None
        try:
            previous_security = excel.AutomationSecurity
            excel.AutomationSecurity = 3  # msoAutomationSecurityForceDisable
        except Exception:
            previous_security = None
        try:
            return excel.Workbooks.Open(path, **kwargs)
        finally:
            if previous_security is not None:
                try:
                    excel.AutomationSecurity = previous_security
                except Exception:
                    pass

    def _read_excel_values(self, workbook: Any) -> tuple[tuple[Any, ...], ...]:
        try:
            sheet = workbook.Worksheets(1)
            used_range = sheet.UsedRange
            source_rows = int(used_range.Rows.Count)
            source_columns = int(used_range.Columns.Count)
            if source_rows > self.MAX_SOURCE_ROWS:
                raise ExcelImportError(
                    "다운로드 Excel 파일의 첫 번째 시트 사용 범위가 "
                    f"안전 처리 한도 {self.MAX_SOURCE_ROWS:,}행을 초과합니다. "
                    "기존 값을 지우지 않았습니다."
                )
            if source_columns > self.MAX_COLUMNS:
                raise ExcelImportError(
                    "다운로드 Excel 파일의 첫 번째 시트 사용 범위가 대상 범위보다 큽니다. "
                    "기존 값을 지우지 않았습니다.\n"
                    f"다운로드 사용 범위: {source_rows}행 × {source_columns}열 / "
                    f"대상: {self.MAX_SOURCE_ROWS}행 × {self.MAX_COLUMNS}열"
                )
            values = self._to_matrix(used_range.Value2)
        except ExcelImportError:
            raise
        except Exception as exc:
            raise ExcelImportError(f"다운로드 Excel 파일의 첫 번째 시트를 읽지 못했습니다.\n{exc}") from exc
        return self._trim_empty_edges(values)

    def _read_delimited_values(self, source_path: Path) -> tuple[tuple[str, ...], ...]:
        with source_path.open("rb") as binary_handle:
            prefix = binary_handle.read(8192)
        last_unicode_error: UnicodeError | None = None
        for encoding in self._candidate_text_encodings(prefix):
            try:
                return self._read_delimited_values_with_encoding(source_path, encoding)
            except UnicodeError as exc:
                # An ASCII-only prefix is valid UTF-8 as well as CP949/EUC-KR.
                # Chrome downloads can contain the first Korean byte after the
                # 8 KiB probe, so retry the complete file with the remaining
                # compatible encodings before reporting it as unreadable.
                last_unicode_error = exc
        raise ExcelImportError("다운로드 텍스트 파일의 문자 인코딩을 읽지 못했습니다.") from last_unicode_error

    def _read_delimited_values_with_encoding(
        self,
        source_path: Path,
        encoding: str,
    ) -> tuple[tuple[str, ...], ...]:
        try:
            with source_path.open("r", encoding=encoding, newline="") as text_handle:
                sample = text_handle.read(8192)
                if self._looks_like_html(sample):
                    raise ExcelImportError(
                        "다운로드 파일이 표 데이터가 아니라 HTML 페이지입니다. 로그인 상태와 다운로드 결과를 확인해 주세요."
                    )
                delimiter = self._detect_delimiter(sample, source_path.suffix.lower())
                text_handle.seek(0)
                rows: list[tuple[str, ...]] = []
                for row_number, row in enumerate(
                    csv.reader(text_handle, delimiter=delimiter),
                    start=1,
                ):
                    if row_number > self.MAX_SOURCE_ROWS:
                        raise ExcelImportError(
                            f"다운로드 데이터가 안전 처리 한도 {self.MAX_SOURCE_ROWS:,}행을 초과합니다. "
                            "기존 값을 지우지 않았습니다."
                        )
                    if len(row) > self.MAX_COLUMNS:
                        raise ExcelImportError(
                            f"다운로드 데이터가 {self.MAX_COLUMNS}열을 초과합니다. 기존 값을 지우지 않았습니다."
                        )
                    rows.append(tuple(row))
        except UnicodeError:
            raise
        except csv.Error as exc:
            raise ExcelImportError(f"다운로드 표 데이터의 구분 형식이 올바르지 않습니다.\n{exc}") from exc
        return self._trim_empty_edges(rows)

    @staticmethod
    def _decode_text(raw: bytes) -> str:
        encoding = MilkrunExcelImporter._detect_text_encoding(raw)
        return raw.decode(encoding)

    @staticmethod
    def _detect_text_encoding(raw: bytes) -> str:
        candidates = MilkrunExcelImporter._candidate_text_encodings(raw)
        if candidates:
            return candidates[0]
        raise ExcelImportError("다운로드 텍스트 파일의 문자 인코딩을 확인할 수 없습니다.")

    @staticmethod
    def _candidate_text_encodings(raw: bytes) -> tuple[str, ...]:
        if raw.startswith((b"\xff\xfe", b"\xfe\xff")):
            return ("utf-16",)
        if raw.startswith(b"\xef\xbb\xbf"):
            return ("utf-8-sig",)
        candidates: list[str] = []
        for encoding in ("utf-8-sig", "cp949", "euc-kr"):
            for trim_count in range(0, min(4, len(raw)) + 1):
                sample = raw if trim_count == 0 else raw[:-trim_count]
                try:
                    sample.decode(encoding)
                    candidates.append(encoding)
                    break
                except UnicodeDecodeError:
                    continue
        return tuple(candidates)

    @classmethod
    def _classify_source(cls, source_path: Path) -> str:
        with source_path.open("rb") as handle:
            raw = handle.read(8192)
        extension = source_path.suffix.lower()
        is_zip_excel = raw.startswith(b"PK\x03\x04")
        is_ole_excel = raw.startswith(b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1")
        if extension in {".xlsx", ".xlsm", ".xlsb"} and is_zip_excel:
            return "excel"
        if extension == ".xls" and is_ole_excel:
            return "excel"
        if extension in cls.SOURCE_EXCEL_EXTENSIONS:
            try:
                text = cls._decode_text(raw)
            except ExcelImportError:
                text = ""
            if cls._looks_like_html(text):
                raise ExcelImportError(
                    "다운로드된 Excel 파일이 실제로는 HTML 페이지입니다. 로그인 상태와 다운로드 결과를 확인해 주세요."
                )
            raise ExcelImportError("다운로드된 Excel 파일이 손상되었거나 지원하지 않는 형식입니다.")

        if extension not in cls.SOURCE_TEXT_EXTENSIONS:
            raise ExcelImportError(f"다운로드 파일 형식을 읽을 수 없습니다: {extension or '(확장자 없음)'}")
        text = cls._decode_text(raw)
        if cls._looks_like_html(text):
            raise ExcelImportError(
                "다운로드 파일이 표 데이터가 아니라 HTML 페이지입니다. 로그인 상태와 다운로드 결과를 확인해 주세요."
            )
        if "\x00" in text:
            raise ExcelImportError("다운로드 파일이 지원하지 않는 바이너리 형식입니다.")
        return "text"

    @staticmethod
    def _looks_like_html(text: str) -> bool:
        prefix = text.lstrip().lower()[:256]
        return prefix.startswith("<")

    @staticmethod
    def _detect_delimiter(text: str, extension: str) -> str:
        if extension == ".tsv":
            return "\t"
        sample = text[:8192]
        try:
            return csv.Sniffer().sniff(sample, delimiters=",\t;|").delimiter
        except csv.Error:
            return "\t" if sample.count("\t") > sample.count(",") else ","

    @staticmethod
    def _to_matrix(value: Any) -> tuple[tuple[Any, ...], ...]:
        if value is None:
            return ()
        if isinstance(value, tuple):
            if not value:
                return ()
            if isinstance(value[0], tuple):
                return tuple(tuple(row) for row in value)
            return (tuple(value),)
        if isinstance(value, list):
            if not value:
                return ()
            if isinstance(value[0], (list, tuple)):
                return tuple(tuple(row) for row in value)
            return (tuple(value),)
        return ((value,),)

    @staticmethod
    def _cell_has_value(value: Any) -> bool:
        return value is not None and value != ""

    @classmethod
    def _trim_empty_edges(
        cls,
        rows: Sequence[Sequence[Any]],
    ) -> tuple[tuple[Any, ...], ...]:
        matrix = [list(row) for row in rows]
        while matrix and not any(cls._cell_has_value(value) for value in matrix[-1]):
            matrix.pop()
        if not matrix:
            return ()

        last_column = 0
        for row in matrix:
            for index, value in enumerate(row, start=1):
                if cls._cell_has_value(value):
                    last_column = max(last_column, index)
        if last_column == 0:
            return ()
        return tuple(tuple(row[:last_column]) for row in matrix)

    @staticmethod
    def _rectangularize(
        rows: Iterable[Sequence[Any]],
        columns: int,
    ) -> tuple[tuple[Any, ...], ...]:
        return tuple(tuple(row) + (None,) * (columns - len(row)) for row in rows)

    @classmethod
    def _filter_arrival_date_rows(
        cls,
        rows: Sequence[Sequence[Any]],
        exclude_arrival_date: date,
    ) -> tuple[tuple[tuple[Any, ...], ...], int]:
        if isinstance(exclude_arrival_date, datetime):
            excluded_date = exclude_arrival_date.date()
        elif isinstance(exclude_arrival_date, date):
            excluded_date = exclude_arrival_date
        else:
            raise ExcelImportError("제외할 입고일 설정이 올바르지 않습니다.")

        header = tuple(rows[0])
        if len(header) < 3:
            raise ExcelImportError("다운로드 데이터의 C열에서 '입고일' 헤더를 찾을 수 없습니다.")
        dispatch_header = re.sub(
            r"\s+",
            "",
            str(header[0] or "").lstrip("\ufeff").strip(),
        )
        if dispatch_header != "배차번호":
            raise ExcelImportError(
                "다운로드 데이터의 A열 헤더가 '배차번호'가 아닙니다. 기존 값을 지우지 않았습니다."
            )
        header_text = str(header[2] or "").lstrip("\ufeff").strip()
        if header_text != "입고일":
            raise ExcelImportError("다운로드 데이터의 C열 헤더가 '입고일'이 아닙니다. 기존 값을 지우지 않았습니다.")

        kept_rows: list[tuple[Any, ...]] = [header]
        filtered_rows = 0
        for row_number, source_row in enumerate(rows[1:], start=2):
            row = tuple(source_row)
            if len(row) < 3 or not cls._cell_has_value(row[2]):
                raise ExcelImportError(
                    f"다운로드 데이터 {row_number}행의 C열 입고일이 비어 있습니다. 기존 값을 지우지 않았습니다."
                )
            arrival_date = cls._parse_arrival_date(row[2])
            if arrival_date is None:
                raise ExcelImportError(
                    f"다운로드 데이터 {row_number}행의 C열 입고일 형식을 확인할 수 없습니다. "
                    "기존 값을 지우지 않았습니다."
                )
            if arrival_date == excluded_date:
                filtered_rows += 1
                continue
            if row and cls._cell_has_value(row[0]) and not normalize_dispatch_number(row[0]):
                raise ExcelImportError(
                    f"다운로드 데이터 {row_number}행의 A열 배차번호를 확인할 수 없습니다. "
                    "기존 값을 지우지 않았습니다."
                )
            kept_rows.append(row)

        # A header-only matrix remains a valid replacement. This clears stale
        # target data while preserving the downloaded column structure.
        return tuple(kept_rows), filtered_rows

    @staticmethod
    def _parse_arrival_date(value: Any) -> date | None:
        # Excel automation and openpyxl-style readers may already expose a
        # proper date/datetime. Handle those before considering numeric serials.
        if isinstance(value, datetime):
            return value.date()
        if isinstance(value, date):
            return value
        if isinstance(value, bool):
            return None

        if isinstance(value, (int, float, Decimal)):
            try:
                serial = float(value)
            except (TypeError, ValueError, OverflowError):
                return None
            if not math.isfinite(serial) or serial < 1 or serial > 2_958_465:
                return None
            try:
                # Excel's 1900 date system includes the historic leap-year
                # compatibility offset; 1899-12-30 is the conventional epoch.
                return (datetime(1899, 12, 30) + timedelta(days=serial)).date()
            except (OverflowError, ValueError):
                return None

        text = str(value or "").strip()
        if not text:
            return None

        year_first = re.fullmatch(
            r"(\d{4})\s*([./-])\s*(\d{1,2})\s*\2\s*(\d{1,2})\s*\.?",
            text,
        )
        if year_first:
            year, month, day = (
                int(year_first.group(1)),
                int(year_first.group(3)),
                int(year_first.group(4)),
            )
            try:
                return date(year, month, day)
            except ValueError:
                return None

        month_first = re.fullmatch(
            r"(\d{1,2})\s*/\s*(\d{1,2})\s*/\s*(\d{4})",
            text,
        )
        if month_first:
            month, day, year = (int(part) for part in month_first.groups())
            try:
                return date(year, month, day)
            except ValueError:
                return None
        return None

    def _save_workbook(
        self,
        excel: Any,
        workbook: Any,
        max_retries: int = 5,
        suppress_display_alerts: bool = False,
        cancel_requested: Callable[[], bool] | None = None,
    ) -> None:
        last_error: Exception | None = None
        for attempt in range(1, max_retries + 1):
            try:
                self._wait_until_excel_ready(
                    excel,
                    cancel_requested=cancel_requested,
                )
                if cancel_requested is not None and cancel_requested():
                    raise ExcelImportCancelled("사용자가 작업을 중지했습니다.")
                if not suppress_display_alerts:
                    workbook.Save()
                else:
                    previous_display_alerts = None
                    try:
                        previous_display_alerts = excel.DisplayAlerts
                        excel.DisplayAlerts = False
                        workbook.Save()
                    finally:
                        if previous_display_alerts is not None:
                            try:
                                excel.DisplayAlerts = previous_display_alerts
                            except Exception:
                                pass
                self.log("연결된 Excel 파일을 저장했습니다.")
                return
            except ExcelImportCancelled:
                raise
            except Exception as exc:
                last_error = exc
                retryable = self._is_excel_busy_error(exc) or any(
                    marker.casefold() in str(exc).casefold()
                    for marker in self._SAVE_RETRY_MARKERS
                )
                if not retryable or attempt >= max_retries:
                    break
                if cancel_requested is not None and cancel_requested():
                    raise ExcelImportCancelled("사용자가 작업을 중지했습니다.") from exc
                self.log(f"Excel 저장을 잠시 기다린 뒤 다시 시도합니다. ({attempt}/{max_retries})")
                self._pump_waiting_com_messages()
                time.sleep(0.5 * attempt)
        raise ExcelImportError(f"연결된 Excel 파일을 저장하지 못했습니다.\n{last_error}")

    def _run_excel_com_operation(
        self,
        excel: Any,
        operation: Callable[[], Any],
        action_name: str,
        *,
        max_retries: int | None = None,
        cancel_requested: Callable[[], bool] | None = None,
    ) -> Any:
        """Run an idempotent Excel COM operation with bounded busy retries."""
        attempts = max_retries or self._COM_OPERATION_MAX_RETRIES
        for attempt in range(1, attempts + 1):
            self._wait_until_excel_ready(
                excel,
                cancel_requested=cancel_requested,
            )
            try:
                return operation()
            except ExcelImportCancelled:
                raise
            except Exception as exc:
                if not self._is_excel_busy_error(exc) or attempt >= attempts:
                    raise
                if cancel_requested is not None and cancel_requested():
                    raise ExcelImportCancelled("사용자가 작업을 중지했습니다.") from exc
                self.log(
                    f"Excel이 다른 작업을 처리 중입니다. "
                    f"{action_name}를 잠시 후 다시 시도합니다. ({attempt}/{attempts})"
                )
                self._pump_waiting_com_messages()
                time.sleep(self._COM_RETRY_BASE_DELAY_SECONDS * attempt)

        raise RuntimeError(f"Excel {action_name} 재시도가 종료되었습니다.")

    def _wait_until_excel_ready(
        self,
        excel: Any,
        *,
        cancel_requested: Callable[[], bool] | None = None,
    ) -> None:
        """Briefly yield while Excel reports that it is busy.

        Some Excel versions do not expose ``Ready`` through all automation
        proxies.  In that case the actual operation remains the source of
        truth and its RPC error is handled by ``_run_excel_com_operation``.
        """
        for check in range(1, self._COM_READY_MAX_CHECKS + 1):
            try:
                if bool(excel.Ready):
                    return
            except Exception as exc:
                if not self._is_excel_busy_error(exc):
                    return
            if check >= self._COM_READY_MAX_CHECKS:
                return
            if cancel_requested is not None and cancel_requested():
                raise ExcelImportCancelled("사용자가 작업을 중지했습니다.")
            self._pump_waiting_com_messages()
            time.sleep(self._COM_READY_BASE_DELAY_SECONDS * check)

    def _pump_waiting_com_messages(self) -> None:
        try:
            pump_messages = getattr(self._pythoncom, "PumpWaitingMessages", None)
            if callable(pump_messages):
                pump_messages()
        except Exception:
            pass

    @classmethod
    def _is_excel_busy_error(cls, exc: BaseException) -> bool:
        hresult_candidates = [getattr(exc, "hresult", None)]
        hresult_candidates.extend(getattr(exc, "args", ()))
        for candidate in hresult_candidates:
            if not isinstance(candidate, int):
                continue
            if candidate in cls._COM_BUSY_HRESULTS:
                return True
            if candidate & 0xFFFFFFFF in {0x80010001, 0x8001010A}:
                return True

        error_text = str(exc).casefold()
        return any(marker.casefold() in error_text for marker in cls._COM_BUSY_RETRY_MARKERS)

    @staticmethod
    def _same_path(left: Path, right: Path) -> bool:
        return os.path.normcase(os.path.abspath(str(left))) == os.path.normcase(os.path.abspath(str(right)))
