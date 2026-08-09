from __future__ import annotations

import tempfile
import unittest
from datetime import date, datetime
from pathlib import Path
from unittest import mock

from Modules.Excel.MilkrunExcelImporter import (
    ExcelImportCancelled,
    ExcelImportError,
    ExcelWorkbookOpenError,
    MilkrunExcelImporter,
)


class FakePythonCom:
    def __init__(self) -> None:
        self.initialized = 0
        self.uninitialized = 0

    def CoInitialize(self) -> None:
        self.initialized += 1

    def CoUninitialize(self) -> None:
        self.uninitialized += 1


class FakeRange:
    def __init__(self, value=None) -> None:
        self.Value2 = value
        self.clear_count = 0
        self.assignments = []

    def ClearContents(self) -> None:
        self.clear_count += 1

    def __setattr__(self, name, value) -> None:
        if name == "Value2" and "assignments" in self.__dict__:
            self.assignments.append(value)
        super().__setattr__(name, value)


class FakeComCallRejected(RuntimeError):
    hresult = -2147418111

    def __init__(self) -> None:
        super().__init__(
            -2147418111,
            "피호출자가 호출을 거부했습니다.",
            None,
            None,
        )


class RejectingRange(FakeRange):
    def __init__(
        self,
        value=None,
        *,
        clear_rejections: int = 0,
        value2_rejections: int = 0,
    ) -> None:
        self.clear_rejections = clear_rejections
        self.value2_rejections = value2_rejections
        self.clear_attempts = 0
        self.value2_attempts = 0
        super().__init__(value)

    def ClearContents(self) -> None:
        self.clear_attempts += 1
        if self.clear_attempts <= self.clear_rejections:
            raise FakeComCallRejected()
        super().ClearContents()

    def __setattr__(self, name, value) -> None:
        if name == "Value2" and "assignments" in self.__dict__:
            self.value2_attempts += 1
            if self.value2_attempts <= self.value2_rejections:
                raise FakeComCallRejected()
        super().__setattr__(name, value)


class FakeCells:
    def __call__(self, row: int, column: int):
        return row, column


class FakeTargetSheet:
    def __init__(self) -> None:
        self.Cells = FakeCells()
        self.clear_range = FakeRange((("old",),))
        self.destination = FakeRange()
        self.destination_coordinates = None

    def Range(self, *args):
        if len(args) == 1 and args[0] == "C1:P1000":
            return self.clear_range
        self.destination_coordinates = args
        return self.destination


class RejectingDestinationRangeSheet(FakeTargetSheet):
    def __init__(self, rejections: int) -> None:
        super().__init__()
        self.destination_range_rejections = rejections
        self.destination_range_attempts = 0

    def Range(self, *args):
        if not (len(args) == 1 and args[0] == "C1:P1000"):
            self.destination_range_attempts += 1
            if self.destination_range_attempts <= self.destination_range_rejections:
                raise FakeComCallRejected()
        return super().Range(*args)


class FakeDimension:
    def __init__(self, count: int) -> None:
        self.Count = count


class FakeUsedRange:
    def __init__(
        self,
        values,
        *,
        row_count: int | None = None,
        column_count: int | None = None,
    ) -> None:
        self._values = values
        self.value2_reads = 0
        matrix = MilkrunExcelImporter._to_matrix(values)
        self.Rows = FakeDimension(row_count if row_count is not None else len(matrix))
        self.Columns = FakeDimension(
            column_count
            if column_count is not None
            else max((len(row) for row in matrix), default=0)
        )

    @property
    def Value2(self):
        self.value2_reads += 1
        return self._values


class FakeSourceSheet:
    def __init__(self, values=None, *, used_range=None) -> None:
        self.UsedRange = used_range or FakeUsedRange(values)


class FakeWorksheets:
    def __init__(
        self,
        *,
        target_sheet=None,
        source_sheet=None,
        target_rejections: int = 0,
    ) -> None:
        self.target_sheet = target_sheet
        self.source_sheet = source_sheet
        self.target_rejections = target_rejections
        self.target_attempts = 0

    def __call__(self, key):
        if key == MilkrunExcelImporter.TARGET_SHEET and self.target_sheet is not None:
            self.target_attempts += 1
            if self.target_attempts <= self.target_rejections:
                raise FakeComCallRejected()
            return self.target_sheet
        if key == 1 and self.source_sheet is not None:
            return self.source_sheet
        raise RuntimeError("sheet not found")


class FakeWorkbook:
    def __init__(
        self,
        full_name: Path,
        worksheets: FakeWorksheets,
        *,
        save_error=None,
        save_errors=(),
        saved: bool = True,
    ) -> None:
        self.FullName = str(full_name)
        self.Worksheets = worksheets
        self.ReadOnly = False
        self.Saved = saved
        self.save_error = save_error
        self.save_errors = list(save_errors)
        self.save_count = 0
        self.close_calls = []

    def Save(self) -> None:
        self.save_count += 1
        if self.save_errors:
            error = self.save_errors.pop(0)
            if error is not None:
                raise error
        if self.save_error is not None:
            raise self.save_error

    def Close(self, SaveChanges=False) -> None:
        self.close_calls.append(SaveChanges)


class BusyStateWorkbook:
    def __init__(
        self,
        worksheets: FakeWorksheets,
        *,
        read_only_rejections: int = 0,
        saved_rejections: int = 0,
    ) -> None:
        self.Worksheets = worksheets
        self.read_only_rejections = read_only_rejections
        self.saved_rejections = saved_rejections
        self.read_only_attempts = 0
        self.saved_attempts = 0

    @property
    def ReadOnly(self):
        self.read_only_attempts += 1
        if self.read_only_attempts <= self.read_only_rejections:
            raise FakeComCallRejected()
        return False

    @property
    def Saved(self):
        self.saved_attempts += 1
        if self.saved_attempts <= self.saved_rejections:
            raise FakeComCallRejected()
        return True


class FakeWorkbooks:
    def __init__(self, open_books=(), open_results=None) -> None:
        self.open_books = list(open_books)
        self.open_results = {str(Path(key).resolve()): value for key, value in (open_results or {}).items()}
        self.open_calls = []

    def __iter__(self):
        return iter(self.open_books)

    def Open(self, path, **kwargs):
        resolved = str(Path(path).resolve())
        self.open_calls.append((resolved, kwargs))
        try:
            return self.open_results[resolved]
        except KeyError as exc:
            raise RuntimeError(f"unexpected open: {resolved}") from exc


class FakeExcel:
    def __init__(self, workbooks: FakeWorkbooks) -> None:
        self.Workbooks = workbooks
        self.Visible = True
        self._display_alerts = True
        self.display_alert_assignments = []
        self.calculate_until_async_queries_done_calls = 0
        self.quit_count = 0

    @property
    def DisplayAlerts(self):
        return self._display_alerts

    @DisplayAlerts.setter
    def DisplayAlerts(self, value) -> None:
        self.display_alert_assignments.append(value)
        self._display_alerts = value

    def CalculateUntilAsyncQueriesDone(self) -> None:
        self.calculate_until_async_queries_done_calls += 1

    def Quit(self) -> None:
        self.quit_count += 1


class FakeComClient:
    def __init__(self, *, active_excel=None, dispatched_excel=None) -> None:
        self.active_excel = active_excel
        self.dispatched_excel = dispatched_excel
        self.dispatch_count = 0

    def GetActiveObject(self, _progid):
        if self.active_excel is None:
            raise RuntimeError("Excel is not running")
        return self.active_excel

    def DispatchEx(self, _progid):
        self.dispatch_count += 1
        if self.dispatched_excel is None:
            raise RuntimeError("Excel is unavailable")
        return self.dispatched_excel


class MilkrunExcelImporterTests(unittest.TestCase):
    def test_skip_target_update_parses_csv_without_loading_excel_com(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            source, target = self._paths(root)
            source.write_text(
                "배차번호,SKU,입고일,수량\n"
                "3369001,100,2026-08-07,1\n"
                "3370492,101,2026-08-08,2\n",
                encoding="utf-8-sig",
            )
            original_target = target.read_bytes()
            logs: list[str] = []
            importer = MilkrunExcelImporter(log=logs.append)

            with mock.patch.object(
                importer,
                "_load_com_modules",
                side_effect=AssertionError("Excel COM must not be loaded"),
            ) as load_com:
                result = importer.import_values(
                    source,
                    target,
                    exclude_arrival_date=date(2026, 8, 7),
                    apply_to_target=False,
                )

            load_com.assert_not_called()
            self.assertFalse(result.target_updated)
            self.assertEqual((result.rows, result.filtered_rows), (2, 1))
            self.assertEqual(result.dispatch_numbers, ("M3370492",))
            self.assertEqual(target.read_bytes(), original_target)
            self.assertTrue(any("시트 반영을 건너뜁니다" in log for log in logs))

    def test_skip_target_update_opens_only_downloaded_excel_read_only(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            source, target = self._paths(root, "download.xlsx")
            source.write_bytes(b"PK\x03\x04placeholder")
            header = [f"열{index}" for index in range(1, 16)]
            header[0] = "배차번호"
            header[2] = "입고일"
            row = [""] * 15
            row[0] = 3370492
            row[1] = "100"
            row[2] = date(2026, 8, 8)
            row[14] = "대상 범위 밖 추가 값"
            source_book = FakeWorkbook(
                source,
                FakeWorksheets(
                    source_sheet=FakeSourceSheet(
                        (
                            tuple(header),
                            tuple(row),
                        )
                    )
                ),
            )
            workbooks = FakeWorkbooks(open_results={source: source_book})
            excel = FakeExcel(workbooks)
            client = FakeComClient(dispatched_excel=excel)
            importer = MilkrunExcelImporter(
                com_client=client,
                pythoncom_module=FakePythonCom(),
            )

            result = importer.import_values(
                source,
                target,
                exclude_arrival_date=date(2026, 8, 7),
                apply_to_target=False,
            )

            self.assertFalse(result.target_updated)
            self.assertEqual(result.dispatch_numbers, ("M3370492",))
            self.assertEqual(result.columns, 15)
            self.assertEqual(client.dispatch_count, 1)
            self.assertEqual(len(workbooks.open_calls), 1)
            self.assertEqual(workbooks.open_calls[0][0], str(source.resolve()))
            self.assertTrue(workbooks.open_calls[0][1]["ReadOnly"])
            self.assertEqual(source_book.close_calls, [False])
            self.assertEqual(excel.quit_count, 1)

    def test_reject_open_target_mode_requests_excel_close_without_mutation(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            _source, target = self._paths(root)
            sheet = FakeTargetSheet()
            target_book = FakeWorkbook(
                target,
                FakeWorksheets(target_sheet=sheet),
            )
            excel = FakeExcel(FakeWorkbooks(open_books=[target_book]))
            client = FakeComClient(active_excel=excel)
            importer = MilkrunExcelImporter(
                com_client=client,
                pythoncom_module=FakePythonCom(),
                reject_open_target=True,
            )

            with self.assertRaisesRegex(
                ExcelWorkbookOpenError,
                "Excel에서 해당 파일을 닫은 뒤",
            ):
                importer.validate_workbook(target)

            self.assertEqual(client.dispatch_count, 0)
            self.assertEqual(sheet.clear_range.clear_count, 0)
            self.assertEqual(target_book.save_count, 0)
            self.assertEqual(target_book.close_calls, [])

    def _paths(self, root: Path, source_name: str = "download.csv") -> tuple[Path, Path]:
        source = root / source_name
        target = root / "입고스케줄관리.xlsx"
        target.write_bytes(b"placeholder")
        return source, target

    @staticmethod
    def _active_importer(target: Path, target_sheet: FakeTargetSheet, *, save_error=None):
        target_book = FakeWorkbook(
            target,
            FakeWorksheets(target_sheet=target_sheet),
            save_error=save_error,
        )
        excel = FakeExcel(FakeWorkbooks(open_books=[target_book]))
        pythoncom = FakePythonCom()
        importer = MilkrunExcelImporter(
            com_client=FakeComClient(active_excel=excel),
            pythoncom_module=pythoncom,
        )
        return importer, target_book, excel, pythoncom

    @staticmethod
    def _excel_source_importer(
        source: Path,
        target: Path,
        target_sheet: FakeTargetSheet,
        source_values,
        *,
        source_range=None,
    ):
        target_book = FakeWorkbook(target, FakeWorksheets(target_sheet=target_sheet))
        source_book = FakeWorkbook(
            source,
            FakeWorksheets(
                source_sheet=FakeSourceSheet(source_values, used_range=source_range)
            ),
        )
        workbooks = FakeWorkbooks(open_books=[target_book], open_results={source: source_book})
        excel = FakeExcel(workbooks)
        importer = MilkrunExcelImporter(
            com_client=FakeComClient(active_excel=excel),
            pythoncom_module=FakePythonCom(),
        )
        return importer, target_book, source_book

    def test_excel_used_range_dimensions_are_rejected_before_value2_read(self) -> None:
        cases = (
            (MilkrunExcelImporter.MAX_SOURCE_ROWS + 1, 4, "10,000행"),
            (2, MilkrunExcelImporter.MAX_COLUMNS + 1, "대상 범위"),
        )
        for row_count, column_count, expected_message in cases:
            with self.subTest(row_count=row_count, column_count=column_count):
                with tempfile.TemporaryDirectory() as temp:
                    root = Path(temp)
                    source, target = self._paths(root, "download.xlsx")
                    source.write_bytes(b"PK\x03\x04placeholder")
                    target_sheet = FakeTargetSheet()
                    source_range = FakeUsedRange(
                        (("센터", "SKU", "입고일", "수량"),),
                        row_count=row_count,
                        column_count=column_count,
                    )
                    importer, target_book, source_book = self._excel_source_importer(
                        source,
                        target,
                        target_sheet,
                        (),
                        source_range=source_range,
                    )

                    with self.assertRaisesRegex(ExcelImportError, expected_message):
                        importer.import_values(source, target)

                    self.assertEqual(source_range.value2_reads, 0)
                    self.assertEqual(target_sheet.clear_range.clear_count, 0)
                    self.assertEqual(target_book.save_count, 0)
                    self.assertEqual(source_book.close_calls, [False])

    def test_excluded_arrival_date_rows_are_removed_before_target_paste(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            source, target = self._paths(root)
            source.write_text(
                "배차번호,SKU,입고일,수량\n"
                "T3369001,100,2026-08-07,1\n"
                ",101,2026. 8. 7.,2\n"
                "3370492,102,2026/08/08,3\n"
                "M3370510,103,08/08/2026,4\n",
                encoding="utf-8-sig",
            )
            original_source = source.read_bytes()
            sheet = FakeTargetSheet()
            importer, target_book, _excel, _pythoncom = self._active_importer(target, sheet)

            result = importer.import_values(
                source,
                target,
                exclude_arrival_date=date(2026, 8, 7),
            )

            self.assertEqual((result.rows, result.columns, result.filtered_rows), (3, 4, 2))
            self.assertEqual(sheet.clear_range.clear_count, 1)
            self.assertEqual(sheet.destination_coordinates, ((1, 3), (3, 6)))
            self.assertEqual(
                sheet.destination.Value2,
                (
                    ("배차번호", "SKU", "입고일", "수량"),
                    ("3370492", "102", "2026/08/08", "3"),
                    ("M3370510", "103", "08/08/2026", "4"),
                ),
            )
            self.assertEqual(result.dispatch_numbers, ("M3370492", "M3370510"))
            self.assertEqual(target_book.save_count, 1)
            self.assertEqual(source.read_bytes(), original_source)

    def test_date_datetime_and_excel_serial_arrival_values_are_supported(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            source, target = self._paths(root, "download.xlsx")
            source.write_bytes(b"PK\x03\x04placeholder")
            original_source = source.read_bytes()
            sheet = FakeTargetSheet()
            excel_serial = (datetime(2026, 8, 7) - datetime(1899, 12, 30)).days
            importer, target_book, source_book = self._excel_source_importer(
                source,
                target,
                sheet,
                (
                    ("배차번호", "SKU", "입고일", "수량"),
                    (3369001, "100", date(2026, 8, 7), 1),
                    (3369002, "101", float(excel_serial), 2),
                    (3370492.0, "102", datetime(2026, 8, 8, 15, 30), 3),
                ),
            )

            result = importer.import_values(
                source,
                target,
                exclude_arrival_date=date(2026, 8, 7),
            )

            self.assertEqual((result.rows, result.filtered_rows), (2, 2))
            self.assertEqual(
                sheet.destination.Value2,
                (
                    ("배차번호", "SKU", "입고일", "수량"),
                    (3370492.0, "102", datetime(2026, 8, 8, 15, 30), 3),
                ),
            )
            self.assertEqual(result.dispatch_numbers, ("M3370492",))
            self.assertEqual(target_book.save_count, 1)
            self.assertEqual(source_book.close_calls, [False])
            self.assertEqual(source.read_bytes(), original_source)

    def test_wrong_arrival_date_header_is_rejected_before_target_change(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            source, target = self._paths(root)
            source.write_text(
                "배차번호,SKU,날짜,수량\n3370492,100,2026-08-08,1\n",
                encoding="utf-8",
            )
            sheet = FakeTargetSheet()
            importer, target_book, _excel, _pythoncom = self._active_importer(target, sheet)

            with self.assertRaisesRegex(ExcelImportError, "C열 헤더"):
                importer.import_values(
                    source,
                    target,
                    exclude_arrival_date=date(2026, 8, 7),
                )

            self.assertEqual(sheet.clear_range.clear_count, 0)
            self.assertEqual(target_book.save_count, 0)

    def test_wrong_dispatch_header_is_rejected_before_target_change(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            source, target = self._paths(root)
            source.write_text(
                "발주번호,SKU,입고일,수량\n3370492,100,2026-08-08,1\n",
                encoding="utf-8",
            )
            sheet = FakeTargetSheet()
            importer, target_book, _excel, _pythoncom = self._active_importer(target, sheet)

            with self.assertRaisesRegex(ExcelImportError, "A열 헤더"):
                importer.import_values(
                    source,
                    target,
                    exclude_arrival_date=date(2026, 8, 7),
                )

            self.assertEqual(sheet.clear_range.clear_count, 0)
            self.assertEqual(target_book.save_count, 0)

    def test_invalid_or_blank_arrival_date_is_rejected_before_target_change(self) -> None:
        for label, arrival_value in (("blank", ""), ("invalid", "2026-02-30")):
            with self.subTest(label=label), tempfile.TemporaryDirectory() as temp:
                root = Path(temp)
                source, target = self._paths(root)
                source.write_text(
                    f"배차번호,SKU,입고일,수량\n3370492,100,{arrival_value},1\n",
                    encoding="utf-8",
                )
                sheet = FakeTargetSheet()
                importer, target_book, _excel, _pythoncom = self._active_importer(target, sheet)

                with self.assertRaisesRegex(ExcelImportError, "C열 입고일"):
                    importer.import_values(
                        source,
                        target,
                        exclude_arrival_date=date(2026, 8, 7),
                    )

                self.assertEqual(sheet.clear_range.clear_count, 0)
                self.assertEqual(target_book.save_count, 0)

    def test_all_excluded_data_rows_paste_header_only(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            source, target = self._paths(root)
            source.write_text(
                "배차번호,SKU,입고일,수량\n"
                "3369001,100,2026.08.07,1\n"
                "3369002,101,08/07/2026,2\n",
                encoding="utf-8",
            )
            sheet = FakeTargetSheet()
            importer, target_book, _excel, _pythoncom = self._active_importer(target, sheet)

            result = importer.import_values(
                source,
                target,
                exclude_arrival_date=date(2026, 8, 7),
            )

            self.assertEqual((result.rows, result.columns, result.filtered_rows), (1, 4, 2))
            self.assertEqual(sheet.destination_coordinates, ((1, 3), (1, 6)))
            self.assertEqual(
                sheet.destination.Value2,
                (("배차번호", "SKU", "입고일", "수량"),),
            )
            self.assertEqual(result.dispatch_numbers, ())
            self.assertEqual(sheet.clear_range.clear_count, 1)
            self.assertEqual(target_book.save_count, 1)

    def test_yesterday_rows_are_filtered_before_target_row_limit(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            source, target = self._paths(root)
            lines = [
                "배차번호,SKU,입고일,수량",
                "3369001,1,2026-08-07,1",
                "3369002,2,2026-08-07,1",
            ]
            lines.extend(
                f"{3370000 + index},{1000 + index},2026-08-08,1" for index in range(999)
            )
            source.write_text("\n".join(lines) + "\n", encoding="utf-8")
            sheet = FakeTargetSheet()
            importer, target_book, _excel, _pythoncom = self._active_importer(target, sheet)

            result = importer.import_values(
                source,
                target,
                exclude_arrival_date=date(2026, 8, 7),
            )

            self.assertEqual((result.rows, result.filtered_rows), (1000, 2))
            self.assertEqual(sheet.clear_range.clear_count, 1)
            self.assertEqual(target_book.save_count, 1)

    def test_csv_values_clear_fixed_range_and_paste_at_c1(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            source, target = self._paths(root)
            source.write_text('센터,수량,메모\n안산2,3,"첫째 줄\n둘째 줄"\n', encoding="utf-8-sig")
            sheet = FakeTargetSheet()
            importer, target_book, excel, pythoncom = self._active_importer(target, sheet)

            result = importer.import_values(source, target)

            self.assertEqual((result.rows, result.columns), (2, 3))
            self.assertEqual(sheet.clear_range.clear_count, 1)
            self.assertEqual(sheet.destination_coordinates, ((1, 3), (2, 5)))
            self.assertEqual(
                sheet.destination.Value2,
                (("센터", "수량", "메모"), ("안산2", "3", "첫째 줄\r\n둘째 줄")),
            )
            self.assertEqual(target_book.save_count, 1)
            self.assertEqual(target_book.close_calls, [])
            self.assertEqual(excel.quit_count, 0)
            self.assertEqual((pythoncom.initialized, pythoncom.uninitialized), (1, 1))

    def test_import_result_contains_unique_a_column_dispatch_numbers(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            source, target = self._paths(root)
            rows = [
                [f"열{index}" for index in range(1, 15)],
                ["3370492"] + [""] * 13,
                [3370492.0] + [""] * 13,
                ["M3370510"] + [""] * 13,
            ]
            source.write_text(
                "\n".join(",".join(str(value) for value in row) for row in rows),
                encoding="utf-8-sig",
            )
            sheet = FakeTargetSheet()
            importer, _target_book, _excel, _pythoncom = self._active_importer(target, sheet)

            result = importer.import_values(source, target)

            self.assertEqual(result.dispatch_numbers, ("M3370492", "M3370510"))

    def test_invalid_today_dispatch_number_is_rejected_before_target_change(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            source, target = self._paths(root)
            source.write_text(
                "배차번호,SKU,입고일,수량\n"
                "T3370492,100,2026-08-08,1\n",
                encoding="utf-8",
            )
            sheet = FakeTargetSheet()
            importer, target_book, _excel, _pythoncom = self._active_importer(target, sheet)

            with self.assertRaisesRegex(ExcelImportError, "A열 배차번호"):
                importer.import_values(
                    source,
                    target,
                    exclude_arrival_date=date(2026, 8, 7),
                )

            self.assertEqual(sheet.clear_range.clear_count, 0)
            self.assertEqual(target_book.save_count, 0)

    def test_blank_dispatch_continuation_row_keeps_first_group_dispatch(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            source, target = self._paths(root)
            source.write_text(
                "배차번호,SKU,입고일,수량\n"
                "3370492,100,2026-08-08,1\n"
                ",101,2026-08-08,2\n",
                encoding="utf-8",
            )
            sheet = FakeTargetSheet()
            importer, target_book, _excel, _pythoncom = self._active_importer(target, sheet)

            result = importer.import_values(
                source,
                target,
                exclude_arrival_date=date(2026, 8, 7),
            )

            self.assertEqual(result.dispatch_numbers, ("M3370492",))
            self.assertEqual(sheet.clear_range.clear_count, 1)
            self.assertEqual(target_book.save_count, 1)

    def test_today_rows_without_dispatch_are_rejected_before_target_change(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            source, target = self._paths(root)
            source.write_text(
                "배차번호,SKU,입고일,수량\n"
                ",100,2026-08-08,1\n",
                encoding="utf-8",
            )
            sheet = FakeTargetSheet()
            importer, target_book, _excel, _pythoncom = self._active_importer(target, sheet)

            with self.assertRaisesRegex(ExcelImportError, "배차번호를 찾지 못했습니다"):
                importer.import_values(
                    source,
                    target,
                    exclude_arrival_date=date(2026, 8, 7),
                )

            self.assertEqual(sheet.clear_range.clear_count, 0)
            self.assertEqual(target_book.save_count, 0)

    def test_formula_like_text_is_copied_as_source_value(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            source, target = self._paths(root)
            source.write_text("name,value\nformula,=WEBSERVICE(\"https://example.test\")\n", encoding="utf-8")
            sheet = FakeTargetSheet()
            importer, _target_book, _excel, _pythoncom = self._active_importer(target, sheet)

            importer.import_values(source, target)

            self.assertEqual(
                sheet.destination.Value2[1][1],
                "=WEBSERVICE(\"https://example.test\")",
            )

    def test_oversized_source_does_not_clear_target(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            source, target = self._paths(root)
            source.write_text(",".join(f"c{index}" for index in range(15)), encoding="utf-8")
            sheet = FakeTargetSheet()
            importer, target_book, _excel, _pythoncom = self._active_importer(target, sheet)

            with self.assertRaisesRegex(ExcelImportError, "14열을 초과"):
                importer.import_values(source, target)

            self.assertEqual(sheet.clear_range.clear_count, 0)
            self.assertEqual(target_book.save_count, 0)

    def test_excel_source_uses_displayed_values_not_formulas_or_formats(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            source, target = self._paths(root, "download.xlsx")
            source.write_bytes(b"PK\x03\x04placeholder")
            sheet = FakeTargetSheet()
            target_book = FakeWorkbook(target, FakeWorksheets(target_sheet=sheet))
            source_book = FakeWorkbook(
                source,
                FakeWorksheets(source_sheet=FakeSourceSheet((("header", 2), ("value", 7), (None, None)))),
            )
            workbooks = FakeWorkbooks(open_books=[target_book], open_results={source: source_book})
            excel = FakeExcel(workbooks)
            importer = MilkrunExcelImporter(
                com_client=FakeComClient(active_excel=excel),
                pythoncom_module=FakePythonCom(),
            )

            result = importer.import_values(source, target)

            self.assertEqual((result.rows, result.columns), (2, 2))
            self.assertEqual(sheet.destination.Value2, (("header", 2), ("value", 7)))
            self.assertEqual(source_book.close_calls, [False])

    def test_html_disguised_as_excel_is_rejected_before_target_change(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            source, target = self._paths(root, "download.xlsx")
            source.write_text("<!doctype html><html><body>login</body></html>", encoding="utf-8")
            sheet = FakeTargetSheet()
            importer, target_book, _excel, pythoncom = self._active_importer(target, sheet)

            with self.assertRaisesRegex(ExcelImportError, "HTML"):
                importer.import_values(source, target)

            self.assertEqual(sheet.clear_range.clear_count, 0)
            self.assertEqual(target_book.save_count, 0)
            self.assertEqual(pythoncom.initialized, 0)

    def test_unknown_text_extension_is_rejected(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            source, target = self._paths(root, "unrelated.log")
            source.write_text("name,value\nitem,1\n", encoding="utf-8")
            sheet = FakeTargetSheet()
            importer, target_book, _excel, pythoncom = self._active_importer(target, sheet)

            with self.assertRaisesRegex(ExcelImportError, "지원 형식"):
                importer.import_values(source, target)

            self.assertEqual(sheet.clear_range.clear_count, 0)
            self.assertEqual(target_book.save_count, 0)
            self.assertEqual(pythoncom.initialized, 0)

    def test_source_size_limit_is_checked_before_com(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            source, target = self._paths(root)
            source.write_text("name,value\nitem,1\n", encoding="utf-8")
            sheet = FakeTargetSheet()
            importer, _target_book, _excel, pythoncom = self._active_importer(target, sheet)

            with (
                mock.patch.object(MilkrunExcelImporter, "MAX_SOURCE_BYTES", 3),
                self.assertRaisesRegex(ExcelImportError, "크기"),
            ):
                importer.import_values(source, target)

            self.assertEqual(pythoncom.initialized, 0)

    def test_utf16_tsv_is_supported(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            source, target = self._paths(root, "download.tsv")
            source.write_bytes("이름\t수량\r\n상품\t2\r\n".encode("utf-16"))
            sheet = FakeTargetSheet()
            importer, _target_book, _excel, _pythoncom = self._active_importer(target, sheet)

            result = importer.import_values(source, target)

            self.assertEqual((result.rows, result.columns), (2, 2))
            self.assertEqual(sheet.destination.Value2, (("이름", "수량"), ("상품", "2")))

    def test_cp949_is_retried_when_first_8k_is_ascii_only(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            source, target = self._paths(root)
            first_value = "a" * 8200
            source.write_bytes(f"{first_value},value\n상품,2\n".encode("cp949"))
            sheet = FakeTargetSheet()
            importer, _target_book, _excel, _pythoncom = self._active_importer(target, sheet)

            result = importer.import_values(source, target)

            self.assertEqual((result.rows, result.columns), (2, 2))
            self.assertEqual(sheet.destination.Value2[1], ("상품", "2"))

    def test_live_excel_global_alerts_and_async_queries_are_untouched(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            source, target = self._paths(root)
            source.write_text("a,b\n1,2\n", encoding="utf-8")
            sheet = FakeTargetSheet()
            importer, target_book, excel, _pythoncom = self._active_importer(target, sheet)

            importer.import_values(source, target)

            self.assertEqual(target_book.save_count, 1)
            self.assertEqual(excel.display_alert_assignments, [])
            self.assertTrue(excel.DisplayAlerts)
            self.assertEqual(excel.calculate_until_async_queries_done_calls, 0)

    def test_hidden_excel_opened_by_importer_is_closed_without_resaving(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            source, target = self._paths(root)
            source.write_text("a,b\n1,2\n", encoding="utf-8")
            sheet = FakeTargetSheet()
            target_book = FakeWorkbook(target, FakeWorksheets(target_sheet=sheet))
            workbooks = FakeWorkbooks(open_results={target: target_book})
            excel = FakeExcel(workbooks)
            client = FakeComClient(dispatched_excel=excel)
            importer = MilkrunExcelImporter(
                com_client=client,
                pythoncom_module=FakePythonCom(),
            )

            importer.import_values(source, target)

            self.assertEqual(client.dispatch_count, 1)
            self.assertEqual(target_book.save_count, 1)
            self.assertEqual(target_book.close_calls, [False])
            self.assertEqual(excel.quit_count, 1)

    def test_save_failure_restores_original_target_values(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            source, target = self._paths(root)
            source.write_text("name,value\nnew,1\n", encoding="utf-8")
            sheet = FakeTargetSheet()
            original = sheet.clear_range.Value2
            importer, target_book, _excel, _pythoncom = self._active_importer(
                target,
                sheet,
                save_error=RuntimeError("disk full"),
            )

            with self.assertRaisesRegex(ExcelImportError, "저장하지 못했습니다"):
                importer.import_values(source, target)

            self.assertEqual(sheet.clear_range.assignments[-1], original)
            self.assertEqual(target_book.save_count, 1)

    def test_transient_excel_busy_retries_clear_and_value_assignment(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            source, target = self._paths(root)
            source.write_text("name,value\nnew,1\n", encoding="utf-8")
            sheet = FakeTargetSheet()
            sheet.clear_range = RejectingRange(
                (("old",),),
                clear_rejections=2,
            )
            sheet.destination = RejectingRange(value2_rejections=2)
            importer, target_book, _excel, _pythoncom = self._active_importer(target, sheet)
            log_messages = []
            importer.log = log_messages.append

            with mock.patch("Modules.Excel.MilkrunExcelImporter.time.sleep"):
                result = importer.import_values(source, target)

            self.assertEqual((result.rows, result.columns), (2, 2))
            self.assertEqual(sheet.clear_range.clear_attempts, 3)
            self.assertEqual(sheet.clear_range.clear_count, 1)
            self.assertEqual(sheet.destination.value2_attempts, 3)
            self.assertEqual(sheet.destination.Value2, (("name", "value"), ("new", "1")))
            self.assertEqual(target_book.save_count, 1)
            self.assertTrue(any("기존 값 지우기" in message for message in log_messages))
            self.assertTrue(any("값 붙여넣기" in message for message in log_messages))

    def test_persistent_excel_busy_before_clear_does_not_attempt_rollback(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            source, target = self._paths(root)
            source.write_text("name,value\nnew,1\n", encoding="utf-8")
            sheet = FakeTargetSheet()
            sheet.clear_range = RejectingRange(
                (("old",),),
                clear_rejections=100,
            )
            importer, target_book, _excel, _pythoncom = self._active_importer(target, sheet)

            with (
                mock.patch("Modules.Excel.MilkrunExcelImporter.time.sleep"),
                self.assertRaisesRegex(ExcelImportError, "Excel 값 붙여넣기에 실패"),
            ):
                importer.import_values(source, target)

            self.assertEqual(
                sheet.clear_range.clear_attempts,
                MilkrunExcelImporter._COM_OPERATION_MAX_RETRIES,
            )
            self.assertEqual(sheet.clear_range.clear_count, 0)
            self.assertEqual(sheet.clear_range.assignments, [])
            self.assertEqual(target_book.save_count, 0)

    def test_transient_excel_busy_retries_destination_range_construction(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            source, target = self._paths(root)
            source.write_text("name,value\nnew,1\n", encoding="utf-8")
            sheet = RejectingDestinationRangeSheet(rejections=2)
            importer, target_book, _excel, _pythoncom = self._active_importer(target, sheet)

            with mock.patch("Modules.Excel.MilkrunExcelImporter.time.sleep"):
                result = importer.import_values(source, target)

            self.assertEqual((result.rows, result.columns), (2, 2))
            self.assertEqual(sheet.destination_range_attempts, 3)
            self.assertEqual(sheet.destination.Value2, (("name", "value"), ("new", "1")))
            self.assertEqual(target_book.save_count, 1)

    def test_persistent_excel_busy_during_assignment_rolls_back(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            source, target = self._paths(root)
            source.write_text("name,value\nnew,1\n", encoding="utf-8")
            sheet = FakeTargetSheet()
            original = sheet.clear_range.Value2
            sheet.destination = RejectingRange(value2_rejections=100)
            importer, target_book, _excel, _pythoncom = self._active_importer(target, sheet)

            with (
                mock.patch("Modules.Excel.MilkrunExcelImporter.time.sleep"),
                self.assertRaisesRegex(ExcelImportError, "값 붙여넣기에 실패"),
            ):
                importer.import_values(source, target)

            self.assertEqual(
                sheet.destination.value2_attempts,
                MilkrunExcelImporter._COM_OPERATION_MAX_RETRIES,
            )
            self.assertEqual(sheet.clear_range.assignments[-1], original)
            self.assertEqual(target_book.save_count, 0)

    def test_persistent_excel_busy_during_rollback_is_reported_explicitly(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            source, target = self._paths(root)
            source.write_text("name,value\nnew,1\n", encoding="utf-8")
            sheet = FakeTargetSheet()
            sheet.clear_range = RejectingRange(
                (("old",),),
                value2_rejections=100,
            )
            sheet.destination = RejectingRange(value2_rejections=100)
            importer, target_book, _excel, _pythoncom = self._active_importer(target, sheet)

            with (
                mock.patch("Modules.Excel.MilkrunExcelImporter.time.sleep"),
                self.assertRaisesRegex(ExcelImportError, "기존 값 복원에도 실패"),
            ):
                importer.import_values(source, target)

            self.assertEqual(
                sheet.clear_range.value2_attempts,
                MilkrunExcelImporter._COM_OPERATION_MAX_RETRIES,
            )
            self.assertEqual(target_book.save_count, 0)

    def test_cancel_during_excel_busy_retry_rolls_back_without_waiting(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            source, target = self._paths(root)
            source.write_text("name,value\nnew,1\n", encoding="utf-8")
            sheet = FakeTargetSheet()
            original = sheet.clear_range.Value2
            sheet.destination = RejectingRange(value2_rejections=100)
            importer, target_book, _excel, _pythoncom = self._active_importer(target, sheet)
            cancel_checks = iter((False, True))

            with (
                mock.patch("Modules.Excel.MilkrunExcelImporter.time.sleep") as sleep_mock,
                self.assertRaises(ExcelImportCancelled),
            ):
                importer.import_values(
                    source,
                    target,
                    cancel_requested=lambda: next(cancel_checks),
                )

            self.assertEqual(sheet.destination.value2_attempts, 1)
            self.assertEqual(sheet.clear_range.assignments[-1], original)
            self.assertEqual(target_book.save_count, 0)
            sleep_mock.assert_not_called()

    def test_cancel_after_assignment_rolls_back_without_saving(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            source, target = self._paths(root)
            source.write_text("name,value\nnew,1\n", encoding="utf-8")
            sheet = FakeTargetSheet()
            original = sheet.clear_range.Value2
            importer, target_book, _excel, _pythoncom = self._active_importer(target, sheet)
            checks = iter((False, True))

            with self.assertRaises(ExcelImportCancelled):
                importer.import_values(
                    source,
                    target,
                    cancel_requested=lambda: next(checks),
                )

            self.assertEqual(target_book.save_count, 0)
            self.assertEqual(sheet.clear_range.assignments[-1], original)

    def test_target_readiness_retries_busy_state_and_sheet_access(self) -> None:
        sheet = FakeTargetSheet()
        worksheets = FakeWorksheets(
            target_sheet=sheet,
            target_rejections=2,
        )
        workbook = BusyStateWorkbook(
            worksheets,
            read_only_rejections=1,
            saved_rejections=1,
        )
        excel = FakeExcel(FakeWorkbooks())
        importer = MilkrunExcelImporter(
            com_client=FakeComClient(active_excel=excel),
            pythoncom_module=FakePythonCom(),
        )

        with mock.patch("Modules.Excel.MilkrunExcelImporter.time.sleep"):
            result = importer._ensure_target_ready(excel, workbook)

        self.assertIs(result, sheet)
        self.assertEqual(workbook.read_only_attempts, 2)
        self.assertEqual(workbook.saved_attempts, 2)
        self.assertEqual(worksheets.target_attempts, 3)

    def test_persistent_busy_sheet_access_is_not_reported_as_missing_sheet(self) -> None:
        sheet = FakeTargetSheet()
        worksheets = FakeWorksheets(
            target_sheet=sheet,
            target_rejections=100,
        )
        workbook = BusyStateWorkbook(worksheets)
        excel = FakeExcel(FakeWorkbooks())
        importer = MilkrunExcelImporter(
            com_client=FakeComClient(active_excel=excel),
            pythoncom_module=FakePythonCom(),
        )

        with (
            mock.patch("Modules.Excel.MilkrunExcelImporter.time.sleep"),
            self.assertRaises(ExcelImportError) as raised,
        ):
            importer._ensure_target_ready(excel, workbook)

        self.assertIn("Excel이 계속 사용 중", str(raised.exception))
        self.assertNotIn("시트가 없습니다", str(raised.exception))

    def test_transient_excel_busy_during_save_is_retried(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            source, target = self._paths(root)
            source.write_text("name,value\nnew,1\n", encoding="utf-8")
            sheet = FakeTargetSheet()
            target_book = FakeWorkbook(
                target,
                FakeWorksheets(target_sheet=sheet),
                save_errors=(FakeComCallRejected(), FakeComCallRejected()),
            )
            excel = FakeExcel(FakeWorkbooks(open_books=[target_book]))
            importer = MilkrunExcelImporter(
                com_client=FakeComClient(active_excel=excel),
                pythoncom_module=FakePythonCom(),
            )

            with mock.patch("Modules.Excel.MilkrunExcelImporter.time.sleep"):
                result = importer.import_values(source, target)

            self.assertEqual((result.rows, result.columns), (2, 2))
            self.assertEqual(target_book.save_count, 3)
            self.assertEqual(sheet.clear_range.assignments, [])

    def test_dirty_open_workbook_is_not_saved_or_changed(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            source, target = self._paths(root)
            source.write_text("name,value\nnew,1\n", encoding="utf-8")
            sheet = FakeTargetSheet()
            target_book = FakeWorkbook(
                target,
                FakeWorksheets(target_sheet=sheet),
                saved=False,
            )
            excel = FakeExcel(FakeWorkbooks(open_books=[target_book]))
            importer = MilkrunExcelImporter(
                com_client=FakeComClient(active_excel=excel),
                pythoncom_module=FakePythonCom(),
            )

            with self.assertRaisesRegex(ExcelImportError, "저장하지 않은 변경사항"):
                importer.import_values(source, target)

            self.assertEqual(sheet.clear_range.clear_count, 0)
            self.assertEqual(target_book.save_count, 0)

    def test_preflight_detects_missing_target_sheet(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            _source, target = self._paths(root)
            target_book = FakeWorkbook(target, FakeWorksheets())
            excel = FakeExcel(FakeWorkbooks(open_books=[target_book]))
            pythoncom = FakePythonCom()
            importer = MilkrunExcelImporter(
                com_client=FakeComClient(active_excel=excel),
                pythoncom_module=pythoncom,
            )

            with self.assertRaisesRegex(ExcelImportError, "시트가 없습니다"):
                importer.validate_workbook(target)

            self.assertEqual((pythoncom.initialized, pythoncom.uninitialized), (1, 1))

    def test_target_filename_must_contain_schedule_management_marker(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            target = Path(temp) / "wrong.xlsx"
            target.write_bytes(b"placeholder")

            with self.assertRaisesRegex(ExcelImportError, "입고스케줄관리"):
                MilkrunExcelImporter.validate_target_path(target)

    def test_cancel_is_honored_before_clear_range_mutation(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            source, target = self._paths(root)
            source.write_text("a,b\n1,2\n", encoding="utf-8")
            sheet = FakeTargetSheet()
            importer, target_book, _excel, _pythoncom = self._active_importer(target, sheet)

            with self.assertRaises(ExcelImportCancelled):
                importer.import_values(source, target, cancel_requested=lambda: True)

            self.assertEqual(sheet.clear_range.clear_count, 0)
            self.assertEqual(target_book.save_count, 0)


if __name__ == "__main__":
    unittest.main()
