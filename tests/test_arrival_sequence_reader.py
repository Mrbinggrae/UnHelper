from __future__ import annotations

import tempfile
import unittest
from decimal import Decimal
from pathlib import Path

from Modules.Excel.ArrivalSequenceReader import (
    ArrivalSequenceEntry,
    ArrivalSequenceReader,
    ArrivalSequenceSnapshot,
    ArrivalSummary,
    ArrivalVehicle,
    BookingFloorAssignment,
    RawBookingAggregate,
    build_floor_target_breakdowns,
    build_arrival_vehicles,
    build_status_pallet_breakdowns,
    normalize_raw_sheet_booking,
    normalize_sequence_booking,
)


class FakePythonCom:
    def __init__(self) -> None:
        self.initialized = 0
        self.uninitialized = 0

    def CoInitialize(self) -> None:
        self.initialized += 1

    def CoUninitialize(self) -> None:
        self.uninitialized += 1


class FakeEnd:
    def __init__(self, row: int) -> None:
        self.Row = row


class FakeCell:
    def __init__(self, row: int) -> None:
        self.row = row

    def End(self, _direction: int) -> FakeEnd:
        return FakeEnd(self.row)


class FakeCells:
    def __init__(self, last_row: int) -> None:
        self.last_row = last_row

    def __call__(self, _row: int, _column: int) -> FakeCell:
        return FakeCell(self.last_row)


class FakeRows:
    Count = 1048576


class FakeRange:
    def __init__(self, values) -> None:
        self.Value2 = values


class FakeSheet:
    def __init__(self, summary_values, detail_values, *, last_row: int) -> None:
        self.Rows = FakeRows()
        self.Cells = FakeCells(last_row)
        self.summary_values = summary_values
        self.detail_values = detail_values
        self.ranges: list[str] = []

    def Range(self, address: str) -> FakeRange:
        self.ranges.append(address)
        if address == "AK8:AT11":
            return FakeRange(self.summary_values)
        return FakeRange(self.detail_values)


class FakeWorksheets:
    def __init__(self, sheet: FakeSheet) -> None:
        truck_values = (("1F", 456),)
        milkrun_values = (("2F", 123),)
        self.sheets = {
            "입차순번": sheet,
            "Raw_트럭": FakeFloorSheet(truck_values),
            "Raw_밀크런": FakeFloorSheet(milkrun_values),
        }

    def __call__(self, key):
        try:
            return self.sheets[key]
        except KeyError as exc:
            raise RuntimeError("sheet not found") from exc


class FakeFloorSheet:
    def __init__(self, values) -> None:
        self.Rows = FakeRows()
        self.Cells = FakeCells(len(values) + 1)
        self.values = values

    def Range(self, _address: str) -> FakeRange:
        return FakeRange(self.values)


class FakeWorkbook:
    def __init__(self, path: Path, sheet: FakeSheet) -> None:
        self.FullName = str(path)
        self.Worksheets = FakeWorksheets(sheet)
        self.close_calls: list[bool] = []

    def Close(self, SaveChanges=False) -> None:
        self.close_calls.append(SaveChanges)


class FakeWorkbooks:
    def __init__(self, open_books=(), open_result=None) -> None:
        self.open_books = list(open_books)
        self.open_result = open_result
        self.open_calls = []

    def __iter__(self):
        return iter(self.open_books)

    def Open(self, path, **kwargs):
        self.open_calls.append((path, kwargs))
        return self.open_result


class FakeExcel:
    def __init__(self, workbooks: FakeWorkbooks) -> None:
        self.Workbooks = workbooks
        self.Ready = True
        self.Visible = True
        self.DisplayAlerts = True
        self.AutomationSecurity = 1
        self.quit_count = 0

    def Quit(self) -> None:
        self.quit_count += 1


class FakeComClient:
    def __init__(self, *, active=None, dispatched=None) -> None:
        self.active = active
        self.dispatched = dispatched
        self.dispatch_count = 0

    def GetActiveObject(self, _progid):
        if self.active is None:
            raise RuntimeError("no active Excel")
        return self.active

    def DispatchEx(self, _progid):
        self.dispatch_count += 1
        return self.dispatched


class ArrivalSequenceReaderTests(unittest.TestCase):
    @staticmethod
    def summary_values():
        return (
            (1, 2, 3, None, 10, 20, None, 4, 5, 6),
            (None,) * 10,
            (7, 8, 9, None, 30, 40, None, 11, 12, 13),
            (14, 15, 16, None, 100, 200, None, 17, 18, 19),
        )

    @staticmethod
    def detail_values():
        first = [None] * 22
        first[0] = "MBN123"
        first[14] = 99
        first[18] = "1F"
        first[21] = 99
        second = [None] * 22
        second[0] = "tbn00456"
        second[14] = 8
        return (tuple(first), tuple(second))

    def test_normalizes_sequence_reservation_prefixes(self) -> None:
        self.assertEqual(normalize_sequence_booking("MBN123"), ("M123", "milkrun"))
        self.assertEqual(normalize_sequence_booking(" mbn000123 "), ("M123", "milkrun"))
        self.assertEqual(normalize_sequence_booking("tbn00123"), ("T123", "truck"))
        self.assertEqual(normalize_sequence_booking("TBN0000456"), ("T456", "truck"))
        self.assertEqual(normalize_sequence_booking("M123"), ("", ""))
        self.assertEqual(normalize_raw_sheet_booking(123, prefix="M"), "M123")
        self.assertEqual(normalize_raw_sheet_booking("00123", prefix="T"), "T123")
        self.assertEqual(normalize_raw_sheet_booking("T123", prefix="T"), "T123")

    def test_reads_live_open_workbook_without_closing_it(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            path = Path(temp) / "SAN2_입고스케줄관리.xlsm"
            path.touch()
            sheet = FakeSheet(self.summary_values(), self.detail_values(), last_row=19)
            workbook = FakeWorkbook(path.resolve(), sheet)
            excel = FakeExcel(FakeWorkbooks(open_books=[workbook]))
            pythoncom = FakePythonCom()
            reader = ArrivalSequenceReader(
                com_client=FakeComClient(active=excel),
                pythoncom_module=pythoncom,
            )

            result = reader.read(path)

            self.assertEqual(result.summary.floor_targets[2], ("100", "200"))
            self.assertEqual([entry.booking_key for entry in result.entries], ["M123", "T456"])
            self.assertEqual(
                [(entry.booking_key, entry.floor) for entry in result.floor_assignments],
                [("T456", "1F"), ("M123", "2F")],
            )
            self.assertEqual(workbook.close_calls, [])
            self.assertEqual(excel.quit_count, 0)
            self.assertEqual(pythoncom.initialized, 1)
            self.assertEqual(pythoncom.uninitialized, 1)

    def test_opens_closed_workbook_read_only_and_closes_owned_excel(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            path = Path(temp) / "SAN2_입고스케줄관리.xlsm"
            path.touch()
            sheet = FakeSheet(self.summary_values(), self.detail_values(), last_row=19)
            workbook = FakeWorkbook(path.resolve(), sheet)
            excel = FakeExcel(FakeWorkbooks(open_result=workbook))
            reader = ArrivalSequenceReader(
                com_client=FakeComClient(dispatched=excel),
                pythoncom_module=FakePythonCom(),
            )

            reader.read(path)

            self.assertEqual(excel.Workbooks.open_calls[0][1]["ReadOnly"], True)
            self.assertEqual(workbook.close_calls, [False])
            self.assertEqual(excel.quit_count, 1)

    def test_current_raw_wins_and_previous_ap_aw_is_counted_once(self) -> None:
        snapshot = ArrivalSequenceSnapshot(
            workbook=Path("sample.xlsm"),
            sheet_name="입차순번",
            refreshed_at=__import__("datetime").datetime.now(),
            summary=ArrivalSummary((), (), ()),
            entries=(
                ArrivalSequenceEntry(18, "MBN123", "M123", "milkrun", "1F", 99, 99),
                ArrivalSequenceEntry(19, "tbn00456", "T456", "truck", "", 8, 8),
                ArrivalSequenceEntry(20, "MBN789", "M789", "milkrun", "2F", 7, 9),
            ),
        )
        raw = {
            "M123": RawBookingAggregate(
                "M123",
                ("거래처",),
                Decimal("3"),
                (("경량", Decimal("2")), ("고단", Decimal("1"))),
            )
        }

        vehicles = build_arrival_vehicles(snapshot, raw)

        self.assertEqual(vehicles[0].period, "금일")
        self.assertEqual(vehicles[0].pallet_count, Decimal("3"))
        self.assertEqual(vehicles[0].categories["고단"], Decimal("1"))
        self.assertEqual(vehicles[1].period, "전일")
        self.assertEqual(vehicles[1].pallet_count, Decimal("8"))
        self.assertEqual(vehicles[1].status, "외부대기")
        self.assertEqual(vehicles[2].pallet_count, Decimal("7"))
        self.assertIn("확인 필요", vehicles[2].note)

    def test_floor_target_breakdown_uses_excel_floor_and_app_raw_prefixes(self) -> None:
        snapshot = ArrivalSequenceSnapshot(
            workbook=Path("sample.xlsm"),
            sheet_name="입차순번",
            refreshed_at=__import__("datetime").datetime.now(),
            summary=ArrivalSummary((), (), ()),
            entries=(),
            floor_assignments=(
                BookingFloorAssignment("T123", "truck", "1F", "Raw_트럭", 2),
                BookingFloorAssignment("M456", "milkrun", "1F", "Raw_밀크런", 2),
                BookingFloorAssignment("T789", "truck", "2F", "Raw_트럭", 3),
            ),
        )
        raw = {
            "T123": RawBookingAggregate(
                "T123", ("트럭",), Decimal("2"), (("중량", Decimal("2")),)
            ),
            "M456": RawBookingAggregate(
                "M456", ("밀크런",), Decimal("3"), (("경량", Decimal("3")),)
            ),
            "T789": RawBookingAggregate(
                "T789", ("트럭2",), Decimal("4"), (("고단", Decimal("4")),)
            ),
        }

        first, second = build_floor_target_breakdowns(snapshot, raw)

        self.assertEqual((first.floor, first.truck_count, first.milkrun_count), ("1F", 1, 1))
        self.assertEqual(first.pallet_count, Decimal("5"))
        self.assertEqual(first.categories["경량"], Decimal("3"))
        self.assertEqual(first.categories["중량"], Decimal("2"))
        self.assertEqual((second.floor, second.truck_count), ("2F", 1))
        self.assertEqual(second.categories["고단"], Decimal("4"))
        self.assertEqual(first.unassigned_raw_bookings, ())

    def test_status_breakdowns_keep_vehicle_counts_separate_from_pallet_details(self) -> None:
        vehicles = (
            ArrivalVehicle(
                18,
                "M123",
                "milkrun",
                "거래처 A",
                "금일",
                "외부대기",
                "1F",
                Decimal("3"),
                (("경량", Decimal("2")), ("고단", Decimal("1"))),
            ),
            ArrivalVehicle(
                19,
                "T456",
                "truck",
                "거래처 B",
                "금일",
                "출차",
                "2F",
                Decimal("4"),
                (("중량", Decimal("4")),),
            ),
            ArrivalVehicle(
                20,
                "M789",
                "milkrun",
                "",
                "전일",
                "외부대기",
                "",
                Decimal("7"),
                (),
            ),
        )

        waiting = build_status_pallet_breakdowns(vehicles, status="외부대기")
        departure = build_status_pallet_breakdowns(vehicles, status="출차")

        self.assertEqual(waiting[0].label, "1F")
        self.assertEqual(waiting[0].pallet_count, Decimal("3"))
        self.assertEqual(waiting[0].categories["경량"], Decimal("2"))
        self.assertEqual(waiting[2].label, "전일자")
        self.assertEqual(waiting[2].pallet_count, Decimal("7"))
        self.assertEqual(departure[1].pallet_count, Decimal("4"))
        self.assertEqual(departure[1].categories["중량"], Decimal("4"))


if __name__ == "__main__":
    unittest.main()
