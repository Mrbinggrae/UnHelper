from __future__ import annotations

import tempfile
import unittest
from pathlib import Path
from unittest import mock

from Modules.Excel.MilkrunExcelImporter import (
    ExcelImportCancelled,
    ExcelImportError,
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


class FakeUsedRange:
    def __init__(self, values) -> None:
        self.Value2 = values


class FakeSourceSheet:
    def __init__(self, values) -> None:
        self.UsedRange = FakeUsedRange(values)


class FakeWorksheets:
    def __init__(self, *, target_sheet=None, source_sheet=None) -> None:
        self.target_sheet = target_sheet
        self.source_sheet = source_sheet

    def __call__(self, key):
        if key == MilkrunExcelImporter.TARGET_SHEET and self.target_sheet is not None:
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
        saved: bool = True,
    ) -> None:
        self.FullName = str(full_name)
        self.Worksheets = worksheets
        self.ReadOnly = False
        self.Saved = saved
        self.save_error = save_error
        self.save_count = 0
        self.close_calls = []

    def Save(self) -> None:
        self.save_count += 1
        if self.save_error is not None:
            raise self.save_error

    def Close(self, SaveChanges=False) -> None:
        self.close_calls.append(SaveChanges)


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

    def test_import_result_contains_unique_p_column_order_numbers(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            source, target = self._paths(root)
            rows = [
                [f"열{index}" for index in range(1, 15)],
                [""] * 13 + ["8789357/8789357"],
                [""] * 13 + ["T8789357"],
                [""] * 13 + ["8827836"],
            ]
            source.write_text(
                "\n".join(",".join(row) for row in rows),
                encoding="utf-8-sig",
            )
            sheet = FakeTargetSheet()
            importer, _target_book, _excel, _pythoncom = self._active_importer(target, sheet)

            result = importer.import_values(source, target)

            self.assertEqual(result.order_numbers, ("8789357", "8827836"))

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
            importer, _target_book, _excel, _pythoncom = self._active_importer(
                target,
                sheet,
                save_error=RuntimeError("disk full"),
            )

            with self.assertRaisesRegex(ExcelImportError, "저장하지 못했습니다"):
                importer.import_values(source, target)

            self.assertEqual(sheet.clear_range.assignments[-1], original)

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
