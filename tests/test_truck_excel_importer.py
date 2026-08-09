from __future__ import annotations

import csv
import tempfile
import unittest
from decimal import Decimal
from pathlib import Path
from unittest import mock

from Modules.Excel import (
    ExcelImportError,
    TruckExcelImporter,
    normalize_truck_reservation_number,
)


class _PythonCom:
    def CoInitialize(self) -> None:
        pass

    def CoUninitialize(self) -> None:
        pass


class _Range:
    def __init__(self, value=None) -> None:
        self.Value2 = value
        self.assignments = []
        self.clear_count = 0

    def __setattr__(self, name, value) -> None:
        if name == "Value2" and "assignments" in self.__dict__:
            self.assignments.append(value)
        super().__setattr__(name, value)

    def ClearContents(self) -> None:
        self.clear_count += 1


class _Cells:
    def __call__(self, row: int, column: int):
        return row, column


class _TruckSheet:
    def __init__(self) -> None:
        self.Cells = _Cells()
        self.clear_range = _Range((("old",),))
        self.destination = _Range()
        self.destination_coordinates = None

    def Range(self, *args):
        if len(args) == 1 and args[0] == "C1:U1000":
            return self.clear_range
        self.destination_coordinates = args
        return self.destination


class _Worksheets:
    def __init__(self, sheet: _TruckSheet) -> None:
        self.sheet = sheet

    def __call__(self, key):
        if key == TruckExcelImporter.TARGET_SHEET:
            return self.sheet
        raise RuntimeError("sheet not found")


class _Workbook:
    def __init__(self, path: Path, sheet: _TruckSheet, *, save_error=None) -> None:
        self.FullName = str(path)
        self.Worksheets = _Worksheets(sheet)
        self.ReadOnly = False
        self.Saved = True
        self.save_error = save_error
        self.save_count = 0

    def Save(self) -> None:
        self.save_count += 1
        if self.save_error is not None:
            raise self.save_error


class _Workbooks:
    def __init__(self, workbook: _Workbook) -> None:
        self.workbook = workbook

    def __iter__(self):
        return iter((self.workbook,))


class _Excel:
    def __init__(self, workbook: _Workbook) -> None:
        self.Workbooks = _Workbooks(workbook)


class _ComClient:
    def __init__(self, excel: _Excel) -> None:
        self.excel = excel

    def GetActiveObject(self, _progid):
        return self.excel


class TruckExcelImporterTests(unittest.TestCase):
    def test_skip_target_update_extracts_truck_metrics_without_excel_com(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            source, target = self._paths(root)
            self._write_csv(
                source,
                (
                    self._headers(),
                    self._row("3370492", "100", "2"),
                ),
            )
            original_target = target.read_bytes()
            importer = TruckExcelImporter()

            with mock.patch.object(
                importer,
                "_load_com_modules",
                side_effect=AssertionError("Excel COM must not be loaded"),
            ) as load_com:
                result = importer.import_values(
                    source,
                    target,
                    apply_to_target=False,
                )

            load_com.assert_not_called()
            self.assertFalse(result.target_updated)
            self.assertEqual(result.dispatch_numbers, ("T3370492",))
            metric = result.metrics_by_reservation["T3370492"]
            self.assertEqual(metric.unit_count, Decimal("100"))
            self.assertEqual(metric.pallet_count, Decimal("2"))
            self.assertEqual(target.read_bytes(), original_target)

    @staticmethod
    def _headers(count: int = 19) -> list[str]:
        headers = [f"열{index}" for index in range(1, max(count, 19) + 1)]
        headers[0] = "예약번호"
        headers[2] = "주문타입"
        headers[4] = "거래처 이름"
        headers[12] = "유닛 수"
        headers[13] = "팔렛트 수"
        return headers[:count]

    @staticmethod
    def _row(
        reservation: object = "",
        units: object = "",
        pallets: object = "",
        *,
        first_value: object = "상품",
        vendor: object = "Test Vendor",
        count: int = 19,
    ) -> list[object]:
        row: list[object] = [""] * count
        row[0] = reservation
        if count >= 3:
            row[2] = first_value
        if count >= 5:
            row[4] = vendor
        if count >= 13:
            row[12] = units
        if count >= 14:
            row[13] = pallets
        return row

    @staticmethod
    def _write_csv(path: Path, rows) -> None:
        with path.open("w", encoding="utf-8-sig", newline="") as handle:
            csv.writer(handle).writerows(rows)

    @staticmethod
    def _paths(root: Path) -> tuple[Path, Path]:
        source = root / "TRUCK_BOOKING_LIST.csv"
        target = root / "SAN2_입고스케줄관리.xlsx"
        target.write_bytes(b"placeholder")
        return source, target

    @staticmethod
    def _importer(target: Path, sheet: _TruckSheet, *, save_error=None):
        workbook = _Workbook(target, sheet, save_error=save_error)
        excel = _Excel(workbook)
        importer = TruckExcelImporter(
            com_client=_ComClient(excel),
            pythoncom_module=_PythonCom(),
        )
        return importer, workbook

    def test_normalizes_truck_reservation_number_without_accepting_other_prefixes(self) -> None:
        self.assertEqual(normalize_truck_reservation_number(" 3370492 "), "T3370492")
        self.assertEqual(normalize_truck_reservation_number("t3370492"), "T3370492")
        self.assertEqual(normalize_truck_reservation_number("3,370,492"), "T3370492")
        self.assertEqual(normalize_truck_reservation_number(3370492.0), "T3370492")
        for invalid in (True, "M3370492", "3370492/1", 3370492.5, "예약번호"):
            with self.subTest(invalid=invalid):
                self.assertEqual(normalize_truck_reservation_number(invalid), "")

    def test_19_columns_clear_truck_range_and_extract_group_metrics(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            source, target = self._paths(root)
            rows = (
                self._headers(),
                self._row("3370492", "", "", first_value="첫 상품"),
                self._row("", "100", "2", first_value="연속 상품"),
                self._row("T3370492", "100.0", "2.0", first_value="중복 예약"),
                self._row("3370493", "45", "3", first_value="다음 예약"),
            )
            self._write_csv(source, rows)
            sheet = _TruckSheet()
            importer, workbook = self._importer(target, sheet)

            result = importer.import_values(source, target)

            self.assertEqual((result.rows, result.columns, result.filtered_rows), (5, 19, 0))
            self.assertEqual(result.sheet_name, "Raw_트럭")
            self.assertEqual(result.dispatch_numbers, ("T3370492", "T3370493"))
            self.assertEqual(sheet.clear_range.clear_count, 1)
            self.assertEqual(sheet.destination_coordinates, ((1, 3), (5, 21)))
            self.assertEqual(sheet.destination.Value2[1][0], "3370492")
            self.assertEqual(sheet.destination.Value2[1][2], "첫 상품")
            self.assertEqual(workbook.save_count, 1)

            first = result.metrics_by_reservation["T3370492"]
            self.assertEqual((first.unit_count, first.pallet_count), (Decimal("100"), Decimal("2")))
            self.assertEqual(first.units_per_pallet, Decimal("50"))
            self.assertEqual(first.source_rows, (2, 3, 4))
            self.assertEqual(first.vendor_name, "Test Vendor")
            second = result.metrics_by_reservation["T3370493"]
            self.assertEqual(second.units_per_pallet, Decimal("15"))

    def test_extracts_vendor_from_e_column_with_blank_continuation_rows(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            source, target = self._paths(root)
            self._write_csv(
                source,
                (
                    self._headers(),
                    self._row("3370492", "", "", vendor="  거래처   A  "),
                    self._row("", "100", "2", vendor=""),
                    self._row("3370493", "45", "3", vendor="거래처 B"),
                ),
            )
            sheet = _TruckSheet()
            importer, _workbook = self._importer(target, sheet)

            result = importer.import_values(source, target)

            self.assertEqual(
                result.metrics_by_reservation["T3370492"].vendor_name,
                "거래처 A",
            )
            self.assertEqual(
                result.metrics_by_reservation["T3370493"].vendor_name,
                "거래처 B",
            )

    def test_conflicting_vendor_for_same_reservation_is_rejected_before_clear(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            source, target = self._paths(root)
            self._write_csv(
                source,
                (
                    self._headers(),
                    self._row("3370492", "100", "2", vendor="거래처 A"),
                    self._row("T3370492", "100", "2", vendor="거래처 B"),
                ),
            )
            sheet = _TruckSheet()
            importer, workbook = self._importer(target, sheet)

            with self.assertRaisesRegex(ExcelImportError, "E열 거래처 이름이 서로 충돌"):
                importer.import_values(source, target)

            self.assertEqual(sheet.clear_range.clear_count, 0)
            self.assertEqual(workbook.save_count, 0)

    def test_conflicting_duplicate_reservation_is_rejected_before_clear(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            source, target = self._paths(root)
            self._write_csv(
                source,
                (
                    self._headers(),
                    self._row("3370492", "100", "2"),
                    self._row("T3370492", "120", "2"),
                ),
            )
            sheet = _TruckSheet()
            importer, workbook = self._importer(target, sheet)

            with self.assertRaisesRegex(ExcelImportError, "서로 충돌"):
                importer.import_values(source, target)

            self.assertEqual(sheet.clear_range.clear_count, 0)
            self.assertEqual(workbook.save_count, 0)

    def test_header_only_download_clears_stale_data_and_returns_no_reservations(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            source, target = self._paths(root)
            self._write_csv(source, (self._headers(),))
            sheet = _TruckSheet()
            importer, workbook = self._importer(target, sheet)

            result = importer.import_values(source, target)

            self.assertEqual((result.rows, result.columns), (1, 19))
            self.assertEqual(result.dispatch_numbers, ())
            self.assertEqual(result.reservation_metrics, ())
            self.assertEqual(sheet.clear_range.clear_count, 1)
            self.assertEqual(sheet.destination_coordinates, ((1, 3), (1, 21)))
            self.assertEqual(workbook.save_count, 1)

    def test_zero_csv_metrics_are_preserved_for_raw_and_daily_detail_lookup(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            source, target = self._paths(root)
            rows = (
                self._headers(),
                self._row("3370492", "0", "2"),
                self._row("3370493", "0", "0"),
            )
            self._write_csv(source, rows)
            sheet = _TruckSheet()
            logs: list[str] = []
            importer, workbook = self._importer(target, sheet)
            importer.log = logs.append

            result = importer.import_values(source, target)

            self.assertEqual(result.dispatch_numbers, ("T3370492", "T3370493"))
            first = result.metrics_by_reservation["T3370492"]
            second = result.metrics_by_reservation["T3370493"]
            self.assertEqual((first.unit_count, first.pallet_count), (Decimal("0"), Decimal("2")))
            self.assertEqual(first.units_per_pallet, Decimal("0"))
            self.assertEqual((second.unit_count, second.pallet_count), (Decimal("0"), Decimal("0")))
            self.assertIsNone(second.units_per_pallet)
            self.assertEqual(sheet.destination.Value2[1][12:14], ("0", "2"))
            self.assertEqual(sheet.destination.Value2[2][12:14], ("0", "0"))
            self.assertEqual(workbook.save_count, 1)
            self.assertEqual(sum("컨테이너 상세 수량" in message for message in logs), 2)

    def test_wrong_or_missing_required_headers_are_rejected_before_clear(self) -> None:
        cases = (
            (0, "발주번호", "A열 예약번호"),
            (12, "박스 수", "M열 유닛 수"),
            (13, "수량", "N열 팔렛트 수"),
        )
        for index, value, expected in cases:
            with self.subTest(index=index), tempfile.TemporaryDirectory() as temp:
                root = Path(temp)
                source, target = self._paths(root)
                headers = self._headers()
                headers[index] = value
                self._write_csv(source, (headers, self._row("3370492", "100", "2")))
                sheet = _TruckSheet()
                importer, workbook = self._importer(target, sheet)

                with self.assertRaisesRegex(ExcelImportError, expected):
                    importer.import_values(source, target)

                self.assertEqual(sheet.clear_range.clear_count, 0)
                self.assertEqual(workbook.save_count, 0)

        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            source, target = self._paths(root)
            self._write_csv(source, (self._headers(13), self._row("3370492", "100", count=13)))
            sheet = _TruckSheet()
            importer, workbook = self._importer(target, sheet)

            with self.assertRaisesRegex(ExcelImportError, "최소 열"):
                importer.import_values(source, target)

            self.assertEqual(sheet.clear_range.clear_count, 0)
            self.assertEqual(workbook.save_count, 0)

    def test_invalid_counts_and_orphan_continuation_are_rejected_before_clear(self) -> None:
        cases = (
            (self._row("3370492", "100", ""), "한쪽만"),
            (self._row("3370492", "-1", "2"), "0 이상의"),
            (self._row("3370492", "NaN", "2"), "유한한"),
            (self._row("", "100", "2"), "직전 예약"),
        )
        for data_row, expected in cases:
            with self.subTest(expected=expected), tempfile.TemporaryDirectory() as temp:
                root = Path(temp)
                source, target = self._paths(root)
                self._write_csv(source, (self._headers(), data_row))
                sheet = _TruckSheet()
                importer, workbook = self._importer(target, sheet)

                with self.assertRaisesRegex(ExcelImportError, expected):
                    importer.import_values(source, target)

                self.assertEqual(sheet.clear_range.clear_count, 0)
                self.assertEqual(workbook.save_count, 0)

    def test_width_and_row_limits_are_rejected_before_clear(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            source, target = self._paths(root)
            self._write_csv(
                source,
                (self._headers(20), self._row("3370492", "100", "2", count=20)),
            )
            sheet = _TruckSheet()
            importer, workbook = self._importer(target, sheet)

            with self.assertRaisesRegex(ExcelImportError, "19열"):
                importer.import_values(source, target)

            self.assertEqual(sheet.clear_range.clear_count, 0)
            self.assertEqual(workbook.save_count, 0)

        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            source, target = self._paths(root)
            rows = [self._headers()]
            rows.extend(self._row("3370492", "100", "2") for _ in range(1000))
            self._write_csv(source, rows)
            sheet = _TruckSheet()
            importer, workbook = self._importer(target, sheet)

            with self.assertRaisesRegex(ExcelImportError, "대상 범위"):
                importer.import_values(source, target)

            self.assertEqual(sheet.clear_range.clear_count, 0)
            self.assertEqual(workbook.save_count, 0)

    def test_save_failure_restores_entire_truck_clear_range(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            source, target = self._paths(root)
            self._write_csv(
                source,
                (self._headers(), self._row("3370492", "100", "2")),
            )
            sheet = _TruckSheet()
            original = sheet.clear_range.Value2
            importer, workbook = self._importer(
                target,
                sheet,
                save_error=RuntimeError("disk full"),
            )

            with self.assertRaisesRegex(ExcelImportError, "저장하지 못했습니다"):
                importer.import_values(source, target)

            self.assertEqual(sheet.clear_range.assignments[-1], original)
            self.assertEqual(workbook.save_count, 1)


if __name__ == "__main__":
    unittest.main()
