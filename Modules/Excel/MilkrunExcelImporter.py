from __future__ import annotations

import csv
import gc
import os
import time
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Callable, Iterable, Sequence

from Modules.Shipments.DailyInbound import extract_order_numbers


LogCallback = Callable[[str], None]


class ExcelImportError(RuntimeError):
    """Raised when downloaded Milkrun values cannot be written safely."""


class ExcelImportCancelled(RuntimeError):
    """Raised before the target workbook mutation starts."""


@dataclass(frozen=True)
class MilkrunExcelImportResult:
    source_file: Path
    target_workbook: Path
    sheet_name: str
    rows: int
    columns: int
    order_numbers: tuple[str, ...]


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
    MAX_COLUMNS = 14
    MAX_SOURCE_BYTES = 50 * 1024 * 1024
    TARGET_FILENAME_MARKER = "입고스케줄관리"
    TARGET_EXTENSIONS = frozenset({".xlsx", ".xlsm", ".xlsb", ".xls"})
    SOURCE_EXCEL_EXTENSIONS = TARGET_EXTENSIONS
    SOURCE_TEXT_EXTENSIONS = frozenset({".csv", ".txt", ".tsv"})
    _SAVE_RETRY_MARKERS = (
        "-2147417846",
        "-2146827284",
        "메시지 필터",
        "문서가 저장되지",
        "Document not saved",
        "Call was rejected by callee",
    )

    def __init__(
        self,
        log: LogCallback | None = None,
        *,
        com_client: Any | None = None,
        pythoncom_module: Any | None = None,
    ):
        self.log = log or (lambda _message: None)
        self._com_client = com_client
        self._pythoncom = pythoncom_module

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
            self._ensure_target_ready(target)
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
            sheet = self._ensure_target_ready(target)
            self.log(f"다운로드 파일의 값 데이터를 읽습니다: {source_path.name}")
            if source_kind == "text":
                values = self._read_delimited_values(source_path)
            else:
                source_book = self._open_source_workbook(excel, source_path)
                values = self._read_excel_values(source_book)

            rows = len(values)
            columns = max((len(row) for row in values), default=0)
            if rows == 0 or columns == 0:
                raise ExcelImportError("다운로드 파일에 붙여넣을 값이 없습니다.")
            if rows > self.MAX_ROWS or columns > self.MAX_COLUMNS:
                raise ExcelImportError(
                    "다운로드 데이터가 대상 범위보다 큽니다. 기존 값을 지우지 않았습니다.\n"
                    f"다운로드: {rows}행 × {columns}열 / 대상: {self.MAX_ROWS}행 × {self.MAX_COLUMNS}열"
                )
            if cancel_requested is not None and cancel_requested():
                raise ExcelImportCancelled("사용자가 작업을 중지했습니다.")

            normalized_values = self._rectangularize(values, columns)
            order_numbers = extract_order_numbers(normalized_values)
            clear_range = sheet.Range(self.CLEAR_RANGE)
            original_attribute = "Value2"
            original_contents = clear_range.Value2
            try:
                original_contents = clear_range.Formula2
                original_attribute = "Formula2"
            except Exception:
                try:
                    original_contents = clear_range.Formula
                    original_attribute = "Formula"
                except Exception:
                    pass
            changed_target = False
            try:
                self.log(f"{self.TARGET_SHEET}!{self.CLEAR_RANGE}의 기존 값만 지웁니다.")
                clear_range.ClearContents()
                changed_target = True
                destination = sheet.Range(
                    sheet.Cells(self.START_ROW, self.START_COLUMN),
                    sheet.Cells(
                        self.START_ROW + rows - 1,
                        self.START_COLUMN + columns - 1,
                    ),
                )
                destination.Value2 = normalized_values
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
                        setattr(clear_range, original_attribute, original_contents)
                    except Exception as exc:
                        rollback_error = exc
                if rollback_error is not None:
                    raise ExcelImportError(
                        "Excel 반영 실패 후 기존 값 복원에도 실패했습니다. 대상 파일을 저장하지 말고 닫은 뒤 다시 열어 주세요.\n"
                        f"반영 오류: {write_error}\n복원 오류: {rollback_error}"
                    ) from write_error
                raise

            return MilkrunExcelImportResult(
                source_file=source_path,
                target_workbook=target_path,
                sheet_name=self.TARGET_SHEET,
                rows=rows,
                columns=columns,
                order_numbers=order_numbers,
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

    def _get_target_sheet(self, workbook: Any) -> Any:
        try:
            return workbook.Worksheets(self.TARGET_SHEET)
        except Exception as exc:
            raise ExcelImportError(
                f"연결된 Excel 파일에 '{self.TARGET_SHEET}' 시트가 없습니다. 기존 값은 변경하지 않았습니다."
            ) from exc

    def _ensure_target_ready(self, workbook: Any) -> Any:
        if bool(getattr(workbook, "ReadOnly", False)):
            raise ExcelImportError(
                "연결된 Excel 파일이 읽기 전용입니다. Excel에서 '편집 사용'으로 연 뒤 다시 시도해 주세요."
            )
        if not bool(getattr(workbook, "Saved", True)):
            raise ExcelImportError(
                "연결된 Excel 파일에 저장하지 않은 변경사항이 있습니다. Excel에서 먼저 저장한 뒤 다시 시도해 주세요."
            )
        return self._get_target_sheet(workbook)

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
            values = self._to_matrix(used_range.Value2)
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
                    if row_number > self.MAX_ROWS:
                        raise ExcelImportError(
                            f"다운로드 데이터가 {self.MAX_ROWS}행을 초과합니다. 기존 값을 지우지 않았습니다."
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
                retryable = any(marker in str(exc) for marker in self._SAVE_RETRY_MARKERS)
                if not retryable or attempt >= max_retries:
                    break
                self.log(f"Excel 저장을 잠시 기다린 뒤 다시 시도합니다. ({attempt}/{max_retries})")
                time.sleep(0.5 * attempt)
        raise ExcelImportError(f"연결된 Excel 파일을 저장하지 못했습니다.\n{last_error}")

    @staticmethod
    def _same_path(left: Path, right: Path) -> bool:
        return os.path.normcase(os.path.abspath(str(left))) == os.path.normcase(os.path.abspath(str(right)))
