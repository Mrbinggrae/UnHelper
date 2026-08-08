from __future__ import annotations

import re
from dataclasses import dataclass
from datetime import date
from decimal import Decimal, InvalidOperation, localcontext
from pathlib import Path
from typing import Any, Callable, cast

from Modules.Excel.MilkrunExcelImporter import (
    ExcelImportError,
    MilkrunExcelImportResult,
    MilkrunExcelImporter,
)


_TRUCK_RESERVATION_PATTERN = re.compile(r"^T?(\d{5,20})$", re.IGNORECASE)


def normalize_truck_reservation_number(value: Any) -> str:
    """Return the canonical ``T<digits>`` key used by daily inbound cards."""

    if value is None or isinstance(value, bool):
        return ""
    if isinstance(value, int):
        candidate = str(value)
    elif isinstance(value, float):
        candidate = str(int(value)) if value.is_integer() else str(value)
    else:
        candidate = str(value)

    candidate = re.sub(r"[\s,]+", "", candidate.strip().upper())
    if candidate.endswith(".0") and re.fullmatch(r"T?\d+", candidate[:-2]):
        candidate = candidate[:-2]
    match = _TRUCK_RESERVATION_PATTERN.fullmatch(candidate)
    return f"T{match.group(1)}" if match else ""


@dataclass(frozen=True)
class TruckReservationMetrics:
    reservation_number: str
    unit_count: Decimal
    pallet_count: Decimal
    source_rows: tuple[int, ...]

    @property
    def units_per_pallet(self) -> Decimal:
        with localcontext() as context:
            context.prec = 28
            return self.unit_count / self.pallet_count


@dataclass(frozen=True)
class TruckExcelImportResult(MilkrunExcelImportResult):
    reservation_metrics: tuple[TruckReservationMetrics, ...] = ()

    @property
    def metrics_by_reservation(self) -> dict[str, TruckReservationMetrics]:
        return {
            metric.reservation_number: metric
            for metric in self.reservation_metrics
        }


class TruckExcelImporter(MilkrunExcelImporter):
    """Replace the Truck RAW area and extract reservation-level quantities."""

    TARGET_SHEET = "Raw_트럭"
    CLEAR_RANGE = "C1:U1000"
    MAX_COLUMNS = 19

    RESERVATION_COLUMN = 3
    UNIT_COUNT_COLUMN = 13
    PALLET_COUNT_COLUMN = 14
    MIN_REQUIRED_COLUMNS = PALLET_COUNT_COLUMN

    RESERVATION_HEADER = "예약번호"
    UNIT_COUNT_HEADERS = frozenset({"유닛수"})
    PALLET_COUNT_HEADERS = frozenset({"팔렛트수", "팔레트수"})

    def import_values(
        self,
        source_file: str | Path,
        target_workbook: str | Path,
        *,
        cancel_requested: Callable[[], bool] | None = None,
    ) -> TruckExcelImportResult:
        result = super().import_values(
            source_file,
            target_workbook,
            cancel_requested=cancel_requested,
            exclude_arrival_date=None,
        )
        return cast(TruckExcelImportResult, result)

    def _extract_import_metadata(
        self,
        values: tuple[tuple[Any, ...], ...],
        *,
        exclude_arrival_date: date | None,
    ) -> tuple[tuple[str, ...], tuple[TruckReservationMetrics, ...]]:
        if exclude_arrival_date is not None:
            raise ExcelImportError("트럭 데이터에는 입고일 제외 필터를 적용할 수 없습니다.")
        if not values:
            raise ExcelImportError("트럭 다운로드 파일에 붙여넣을 값이 없습니다.")

        column_count = len(values[0])
        if column_count < self.MIN_REQUIRED_COLUMNS:
            raise ExcelImportError(
                "트럭 다운로드 데이터에서 예약번호와 유닛/팔렛트 수 열을 찾을 수 없습니다. "
                "기존 값을 지우지 않았습니다.\n"
                f"필요한 최소 열: {self.MIN_REQUIRED_COLUMNS}열 / 다운로드: {column_count}열"
            )

        header = values[0]
        self._validate_header(
            header,
            self.RESERVATION_COLUMN,
            frozenset({self.RESERVATION_HEADER}),
            "C열 예약번호",
        )
        self._validate_header(
            header,
            self.UNIT_COUNT_COLUMN,
            self.UNIT_COUNT_HEADERS,
            "M열 유닛 수",
        )
        self._validate_header(
            header,
            self.PALLET_COUNT_COLUMN,
            self.PALLET_COUNT_HEADERS,
            "N열 팔렛트 수",
        )

        ordered_reservations: list[str] = []
        source_rows_by_reservation: dict[str, list[int]] = {}
        metric_values: dict[str, tuple[Decimal, Decimal]] = {}
        metric_source_row: dict[str, int] = {}
        current_reservation = ""

        for row_number, row in enumerate(values[1:], start=2):
            raw_reservation = row[self.RESERVATION_COLUMN - 1]
            if self._has_import_value(raw_reservation):
                current_reservation = normalize_truck_reservation_number(raw_reservation)
                if not current_reservation:
                    raise ExcelImportError(
                        f"트럭 다운로드 데이터 {row_number}행의 C열 예약번호를 확인할 수 없습니다. "
                        "기존 값을 지우지 않았습니다."
                    )
                if current_reservation not in source_rows_by_reservation:
                    ordered_reservations.append(current_reservation)
                    source_rows_by_reservation[current_reservation] = []
            elif not any(self._has_import_value(value) for value in row):
                continue
            elif not current_reservation:
                raise ExcelImportError(
                    f"트럭 다운로드 데이터 {row_number}행의 C열 예약번호가 비어 있어 "
                    "직전 예약 데이터에 연결할 수 없습니다. 기존 값을 지우지 않았습니다."
                )

            source_rows_by_reservation[current_reservation].append(row_number)
            raw_units = row[self.UNIT_COUNT_COLUMN - 1]
            raw_pallets = row[self.PALLET_COUNT_COLUMN - 1]
            has_units = self._has_import_value(raw_units)
            has_pallets = self._has_import_value(raw_pallets)
            if has_units != has_pallets:
                raise ExcelImportError(
                    f"트럭 다운로드 데이터 {row_number}행은 M열 유닛 수와 N열 팔렛트 수 중 "
                    "한쪽만 입력되어 있습니다. 기존 값을 지우지 않았습니다."
                )
            if not has_units:
                continue

            units = self._positive_decimal(raw_units, "M열 유닛 수", row_number)
            pallets = self._positive_decimal(raw_pallets, "N열 팔렛트 수", row_number)
            candidate = (units, pallets)
            previous = metric_values.get(current_reservation)
            if previous is not None and previous != candidate:
                previous_row = metric_source_row[current_reservation]
                raise ExcelImportError(
                    f"트럭 예약번호 {current_reservation}의 유닛/팔렛트 수가 서로 충돌합니다. "
                    "합산하거나 마지막 값을 선택하지 않았으며 기존 값도 지우지 않았습니다.\n"
                    f"{previous_row}행: {previous[0]} / {previous[1]} · "
                    f"{row_number}행: {units} / {pallets}"
                )
            if previous is None:
                metric_values[current_reservation] = candidate
                metric_source_row[current_reservation] = row_number

        metrics: list[TruckReservationMetrics] = []
        for reservation_number in ordered_reservations:
            pair = metric_values.get(reservation_number)
            if pair is None:
                rows_text = ", ".join(
                    str(row_number)
                    for row_number in source_rows_by_reservation[reservation_number]
                )
                raise ExcelImportError(
                    f"트럭 예약번호 {reservation_number}의 M열 유닛 수와 N열 팔렛트 수를 "
                    f"찾지 못했습니다. 해당 행: {rows_text}. 기존 값을 지우지 않았습니다."
                )
            metrics.append(
                TruckReservationMetrics(
                    reservation_number=reservation_number,
                    unit_count=pair[0],
                    pallet_count=pair[1],
                    source_rows=tuple(source_rows_by_reservation[reservation_number]),
                )
            )

        return tuple(ordered_reservations), tuple(metrics)

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
    ) -> TruckExcelImportResult:
        return TruckExcelImportResult(
            source_file=source_file,
            target_workbook=target_workbook,
            sheet_name=sheet_name,
            rows=rows,
            columns=columns,
            dispatch_numbers=dispatch_numbers,
            filtered_rows=filtered_rows,
            reservation_metrics=tuple(import_metadata or ()),
        )

    @classmethod
    def _validate_header(
        cls,
        row: tuple[Any, ...],
        column: int,
        expected: frozenset[str],
        label: str,
    ) -> None:
        actual = cls._normalize_header(row[column - 1])
        if actual not in expected:
            expected_text = " 또는 ".join(sorted(expected))
            raise ExcelImportError(
                f"트럭 다운로드 데이터의 {label} 헤더가 올바르지 않습니다. "
                f"필요한 헤더: {expected_text} / 확인된 값: {actual or '(빈 값)'}. "
                "기존 값을 지우지 않았습니다."
            )

    @staticmethod
    def _normalize_header(value: Any) -> str:
        return re.sub(r"\s+", "", str(value or "").lstrip("\ufeff").strip())

    @staticmethod
    def _has_import_value(value: Any) -> bool:
        if value is None:
            return False
        if isinstance(value, str):
            return bool(value.strip())
        return True

    @staticmethod
    def _positive_decimal(value: Any, label: str, row_number: int) -> Decimal:
        if isinstance(value, bool):
            raise ExcelImportError(
                f"트럭 다운로드 데이터 {row_number}행의 {label} 값이 숫자가 아닙니다. "
                "기존 값을 지우지 않았습니다."
            )
        candidate = re.sub(r"[\s,]+", "", str(value))
        try:
            result = Decimal(candidate)
        except (InvalidOperation, ValueError) as exc:
            raise ExcelImportError(
                f"트럭 다운로드 데이터 {row_number}행의 {label} 값이 숫자가 아닙니다: {value!r}. "
                "기존 값을 지우지 않았습니다."
            ) from exc
        if not result.is_finite() or result <= 0:
            raise ExcelImportError(
                f"트럭 다운로드 데이터 {row_number}행의 {label} 값은 0보다 큰 유한한 숫자여야 합니다: "
                f"{value!r}. 기존 값을 지우지 않았습니다."
            )
        return result
