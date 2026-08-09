from __future__ import annotations

import gc
import re
from dataclasses import dataclass
from datetime import datetime
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import Any, Callable, Mapping

from Modules.Excel.MilkrunExcelImporter import (
    ExcelImportCancelled,
    ExcelImportError,
    MilkrunExcelImporter,
)


LogCallback = Callable[[str], None]


class ArrivalSequenceError(ExcelImportError):
    """Raised when the linked arrival-sequence sheet cannot be read safely."""


@dataclass(frozen=True)
class ArrivalSummary:
    departure: tuple[tuple[str, str, str], ...]
    outside_waiting: tuple[tuple[str, str, str], ...]
    floor_targets: tuple[tuple[str, str], ...]


@dataclass(frozen=True)
class ArrivalSequenceEntry:
    excel_row: int
    raw_reservation: str
    booking_key: str
    booking_type: str
    unloading_floor: str
    previous_ap_value: Any = None
    previous_aw_value: Any = None


@dataclass(frozen=True)
class ArrivalSequenceSnapshot:
    workbook: Path
    sheet_name: str
    refreshed_at: datetime
    summary: ArrivalSummary
    entries: tuple[ArrivalSequenceEntry, ...]


@dataclass(frozen=True)
class RawBookingAggregate:
    booking_key: str
    vendor_names: tuple[str, ...]
    pallet_count: Decimal
    category_pallets: tuple[tuple[str, Decimal], ...]
    missing_pallet_rows: int = 0

    @property
    def categories(self) -> dict[str, Decimal]:
        return dict(self.category_pallets)


@dataclass(frozen=True)
class ArrivalVehicle:
    excel_row: int
    booking_key: str
    booking_type: str
    vendor_name: str
    period: str
    status: str
    floor: str
    pallet_count: Decimal | None
    category_pallets: tuple[tuple[str, Decimal], ...]
    note: str = ""

    @property
    def categories(self) -> dict[str, Decimal]:
        return dict(self.category_pallets)


_MILKRUN_SEQUENCE_PATTERN = re.compile(r"^MBN\s*0*(\d+)\s*$", re.IGNORECASE)
_TRUCK_SEQUENCE_PATTERN = re.compile(r"^TBN00\s*0*(\d+)\s*$", re.IGNORECASE)
_FLOOR_PATTERN = re.compile(r"\b([12])\s*F\b", re.IGNORECASE)


def normalize_sequence_booking(value: Any) -> tuple[str, str]:
    """Map the sequence sheet's MBN/tbn00 values to RAW M/T keys."""

    candidate = re.sub(r"\s+", "", str(value or "")).strip()
    if not candidate:
        return "", ""
    match = _MILKRUN_SEQUENCE_PATTERN.fullmatch(candidate)
    if match:
        return f"M{match.group(1)}", "milkrun"
    match = _TRUCK_SEQUENCE_PATTERN.fullmatch(candidate)
    if match:
        return f"T{match.group(1)}", "truck"
    return "", ""


def decimal_or_none(value: Any) -> Decimal | None:
    if value is None or isinstance(value, bool):
        return None
    candidate = str(value).strip().replace(",", "")
    if not candidate:
        return None
    try:
        parsed = Decimal(candidate)
    except InvalidOperation:
        return None
    if not parsed.is_finite() or parsed < 0:
        return None
    return parsed


def build_arrival_vehicles(
    snapshot: ArrivalSequenceSnapshot,
    raw_bookings: Mapping[str, RawBookingAggregate],
) -> tuple[ArrivalVehicle, ...]:
    """Combine live Excel status with current-day RAW pallet classifications."""

    vehicles: list[ArrivalVehicle] = []
    for entry in snapshot.entries:
        floor_match = _FLOOR_PATTERN.search(entry.unloading_floor)
        floor = f"{floor_match.group(1)}F" if floor_match else ""
        status = "출차" if entry.unloading_floor.strip() else "외부대기"
        current = raw_bookings.get(entry.booking_key)
        if current is not None:
            note = (
                f"팔렛트 수 미입력 SKU {current.missing_pallet_rows}행"
                if current.missing_pallet_rows
                else ""
            )
            vehicles.append(
                ArrivalVehicle(
                    excel_row=entry.excel_row,
                    booking_key=entry.booking_key,
                    booking_type=entry.booking_type,
                    vendor_name=" · ".join(current.vendor_names),
                    period="금일",
                    status=status,
                    floor=floor,
                    pallet_count=current.pallet_count,
                    category_pallets=current.category_pallets,
                    note=note,
                )
            )
            continue

        ap_value = decimal_or_none(entry.previous_ap_value)
        aw_value = decimal_or_none(entry.previous_aw_value)
        previous_value = ap_value if ap_value is not None else aw_value
        note = ""
        if ap_value is not None and aw_value is not None and ap_value != aw_value:
            note = f"전일 팔렛트 값 확인 필요 (AP {ap_value} / AW {aw_value})"
        elif previous_value is None:
            note = "전일 팔렛트 수 미입력"
        vehicles.append(
            ArrivalVehicle(
                excel_row=entry.excel_row,
                booking_key=entry.booking_key,
                booking_type=entry.booking_type,
                vendor_name="",
                period="전일",
                status=status,
                floor=floor,
                pallet_count=previous_value,
                category_pallets=(),
                note=note,
            )
        )
    return tuple(vehicles)


class ArrivalSequenceReader(MilkrunExcelImporter):
    """Read the linked workbook's ``입차순번`` sheet without changing it."""

    TARGET_SHEET = "입차순번"
    FIRST_DETAIL_ROW = 18
    MAX_DETAIL_ROW = 5000
    _XL_UP = -4162

    def __init__(
        self,
        log: LogCallback | None = None,
        *,
        com_client: Any | None = None,
        pythoncom_module: Any | None = None,
    ) -> None:
        super().__init__(
            log=log,
            com_client=com_client,
            pythoncom_module=pythoncom_module,
            reject_open_target=False,
        )

    def read(
        self,
        workbook_path: str | Path,
        *,
        cancel_requested: Callable[[], bool] | None = None,
    ) -> ArrivalSequenceSnapshot:
        target_path = self.validate_target_path(workbook_path)
        pythoncom, com_client = self._load_com_modules()
        coinitialized = False
        excel = None
        workbook = None
        owns_excel = False
        owns_workbook = False
        try:
            pythoncom.CoInitialize()
            coinitialized = True
            excel, workbook, owns_excel, owns_workbook = self._open_readonly_workbook(
                com_client,
                target_path,
            )
            sheet = self._run_excel_com_operation(
                excel,
                lambda: workbook.Worksheets(self.TARGET_SHEET),
                "입차순번 시트 확인",
                cancel_requested=cancel_requested,
            )
            summary_values = self._run_excel_com_operation(
                excel,
                lambda: sheet.Range("AK8:AT11").Value2,
                "입차순번 요약 값 읽기",
                cancel_requested=cancel_requested,
            )
            last_row = self._last_detail_row(
                excel,
                sheet,
                cancel_requested=cancel_requested,
            )
            detail_values = ()
            if last_row >= self.FIRST_DETAIL_ROW:
                detail_values = self._run_excel_com_operation(
                    excel,
                    lambda: sheet.Range(
                        f"AB{self.FIRST_DETAIL_ROW}:AW{last_row}"
                    ).Value2,
                    "입차순번 차량 값 읽기",
                    cancel_requested=cancel_requested,
                )
            snapshot = ArrivalSequenceSnapshot(
                workbook=target_path,
                sheet_name=self.TARGET_SHEET,
                refreshed_at=datetime.now(),
                summary=self._parse_summary(summary_values),
                entries=self._parse_entries(detail_values),
            )
            self.log(
                f"입차순번 시트에서 차량 {len(snapshot.entries)}건을 읽었습니다."
            )
            return snapshot
        except ExcelImportCancelled:
            raise
        except ArrivalSequenceError:
            raise
        except Exception as exc:
            raise ArrivalSequenceError(
                "연결된 Excel의 '입차순번' 시트를 읽지 못했습니다. "
                "Excel의 편집 중인 셀이나 팝업창을 닫고 다시 시도해 주세요.\n"
                f"{exc}"
            ) from exc
        finally:
            if workbook is not None and owns_workbook:
                try:
                    workbook.Close(SaveChanges=False)
                except Exception:
                    pass
            if excel is not None and owns_excel:
                try:
                    excel.Quit()
                except Exception:
                    pass
            workbook = None
            excel = None
            gc.collect()
            if coinitialized:
                try:
                    pythoncom.CoUninitialize()
                except Exception:
                    pass

    def _open_readonly_workbook(
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
            open_target = self._run_excel_com_operation(
                active_excel,
                lambda: self._find_open_workbook_for_read(active_excel, target_path),
                "열려 있는 Excel 파일 찾기",
            )
            if open_target is not None:
                self.log(f"열려 있는 Excel 파일의 현재 값을 읽습니다: {target_path.name}")
                return active_excel, open_target, False, False

        excel = None
        try:
            excel = com_client.DispatchEx("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False
            workbook = self._workbooks_open_with_macros_disabled(
                excel,
                str(target_path),
                UpdateLinks=0,
                ReadOnly=True,
                IgnoreReadOnlyRecommended=True,
                AddToMru=False,
            )
            self.log(f"닫힌 Excel 파일을 읽기 전용으로 열었습니다: {target_path.name}")
            return excel, workbook, True, True
        except Exception as exc:
            if excel is not None:
                try:
                    excel.Quit()
                except Exception:
                    pass
            raise ArrivalSequenceError(
                f"연결된 Excel 파일을 읽기 전용으로 열지 못했습니다.\n{exc}"
            ) from exc

    def _find_open_workbook_for_read(
        self,
        excel: Any,
        target_path: Path,
    ) -> Any | None:
        for workbook in excel.Workbooks:
            try:
                full_name = Path(str(workbook.FullName))
            except Exception as exc:
                if self._is_excel_busy_error(exc):
                    raise
                continue
            if self._same_path(full_name, target_path):
                return workbook
        return None

    def _last_detail_row(
        self,
        excel: Any,
        sheet: Any,
        *,
        cancel_requested: Callable[[], bool] | None,
    ) -> int:
        def find_last_row() -> int:
            row_count = int(sheet.Rows.Count)
            columns = (28, 42, 46, 49)  # AB, AP, AT, AW
            rows = [int(sheet.Cells(row_count, column).End(self._XL_UP).Row) for column in columns]
            return max(rows, default=0)

        last_row = self._run_excel_com_operation(
            excel,
            find_last_row,
            "입차순번 마지막 행 확인",
            cancel_requested=cancel_requested,
        )
        return min(max(int(last_row), self.FIRST_DETAIL_ROW - 1), self.MAX_DETAIL_ROW)

    @staticmethod
    def _matrix(values: Any) -> tuple[tuple[Any, ...], ...]:
        return MilkrunExcelImporter._to_matrix(values) if values not in (None, ()) else ()

    @classmethod
    def _parse_summary(cls, values: Any) -> ArrivalSummary:
        matrix = cls._matrix(values)

        def at(row: int, column: int) -> str:
            try:
                value = matrix[row][column]
            except IndexError:
                return ""
            if value is None:
                return ""
            if isinstance(value, float) and value.is_integer():
                return str(int(value))
            return str(value).strip()

        return ArrivalSummary(
            departure=(
                (at(0, 0), at(0, 1), at(0, 2)),
                (at(2, 0), at(2, 1), at(2, 2)),
                (at(3, 0), at(3, 1), at(3, 2)),
            ),
            outside_waiting=(
                (at(0, 7), at(0, 8), at(0, 9)),
                (at(2, 7), at(2, 8), at(2, 9)),
                (at(3, 7), at(3, 8), at(3, 9)),
            ),
            floor_targets=(
                (at(0, 4), at(0, 5)),
                (at(2, 4), at(2, 5)),
                (at(3, 4), at(3, 5)),
            ),
        )

    @classmethod
    def _parse_entries(cls, values: Any) -> tuple[ArrivalSequenceEntry, ...]:
        entries: list[ArrivalSequenceEntry] = []
        for offset, row in enumerate(cls._matrix(values)):
            raw_reservation = row[0] if len(row) > 0 else None
            booking_key, booking_type = normalize_sequence_booking(raw_reservation)
            if not booking_key:
                continue
            entries.append(
                ArrivalSequenceEntry(
                    excel_row=cls.FIRST_DETAIL_ROW + offset,
                    raw_reservation=str(raw_reservation or "").strip(),
                    booking_key=booking_key,
                    booking_type=booking_type,
                    previous_ap_value=row[14] if len(row) > 14 else None,
                    unloading_floor=str(row[18] or "").strip() if len(row) > 18 else "",
                    previous_aw_value=row[21] if len(row) > 21 else None,
                )
            )
        return tuple(entries)
