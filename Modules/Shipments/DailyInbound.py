from __future__ import annotations

import re
from dataclasses import dataclass
from decimal import Decimal, InvalidOperation
from typing import Any, Iterable, Sequence


_BOOKING_NUMBER_PATTERN = re.compile(r"^([MT])?(\d{5,20})$", re.IGNORECASE)
_SUPPORTED_BOOKING_PREFIXES = frozenset({"M", "T"})
_VENDOR_CODE_SUFFIX = re.compile(r"\s*\(\s*A[0-9A-Z_-]+\s*\)\s*$", re.IGNORECASE)


@dataclass(frozen=True)
class MilkrunProductRow:
    vendor_name: str
    milkrun_number: str
    pallet_count: str
    box_count: str
    sku_id: str
    sku_name: str
    dispatch_number: str = ""
    order_number: str = ""


def normalize_booking_number(value: Any, *, prefix: str) -> str:
    """Return a source booking number as an exact prefixed schedule key.

    Downloaded source files may contain bare numbers or a key that already has
    the expected prefix.  A key carrying another booking type's prefix is
    rejected so an ``M`` Milkrun and ``T`` Truck booking with the same digits
    can never be treated as the same schedule card.
    """

    expected_prefix = str(prefix or "").strip().upper()
    if expected_prefix not in _SUPPORTED_BOOKING_PREFIXES:
        raise ValueError(f"지원하지 않는 입고 예약 접두사입니다: {prefix!r}")

    if value is None or isinstance(value, bool):
        return ""
    if isinstance(value, int):
        candidate = str(value)
    elif isinstance(value, float):
        candidate = str(int(value)) if value.is_integer() else str(value)
    else:
        candidate = str(value)

    candidate = candidate.strip().upper()
    candidate = re.sub(r"[\s,]+", "", candidate)
    if candidate.endswith(".0") and re.fullmatch(r"[MT]?\d+", candidate[:-2]):
        candidate = candidate[:-2]
    match = _BOOKING_NUMBER_PATTERN.fullmatch(candidate)
    if not match:
        return ""
    source_prefix, digits = match.groups()
    if source_prefix and source_prefix.upper() != expected_prefix:
        return ""
    return f"{expected_prefix}{digits}"


def normalize_dispatch_number(value: Any) -> str:
    """Return the canonical ``M<digits>`` key for a Milkrun source value."""

    return normalize_booking_number(value, prefix="M")


def normalize_truck_reservation_number(value: Any) -> str:
    """Return the canonical ``T<digits>`` key for a Truck source value."""

    return normalize_booking_number(value, prefix="T")


def normalize_booking_card_number(value: Any, *, prefix: str) -> str:
    """Accept only a card label that explicitly carries ``prefix``."""

    expected_prefix = str(prefix or "").strip().upper()
    if expected_prefix not in _SUPPORTED_BOOKING_PREFIXES:
        raise ValueError(f"지원하지 않는 입고 예약 접두사입니다: {prefix!r}")
    if not isinstance(value, str):
        return ""
    candidate = re.sub(r"[\s,]+", "", value.strip().upper())
    if not candidate.startswith(expected_prefix):
        return ""
    return normalize_booking_number(candidate, prefix=expected_prefix)


def normalize_milkrun_card_number(value: Any) -> str:
    """Accept only a schedule-card label that explicitly starts with ``M``."""

    return normalize_booking_card_number(value, prefix="M")


def normalize_truck_card_number(value: Any) -> str:
    """Accept only a schedule-card label that explicitly starts with ``T``."""

    return normalize_booking_card_number(value, prefix="T")


def extract_booking_numbers(
    rows: Iterable[Sequence[Any]],
    *,
    source_column: int,
    prefix: str,
) -> tuple[str, ...]:
    """Read unique booking numbers from a one-based source column."""

    if source_column < 1:
        raise ValueError("source_column must be 1 or greater")

    index = source_column - 1
    seen: set[str] = set()
    result: list[str] = []
    for row in rows:
        if len(row) <= index:
            continue
        booking_number = normalize_booking_number(row[index], prefix=prefix)
        if booking_number and booking_number not in seen:
            seen.add(booking_number)
            result.append(booking_number)
    return tuple(result)


def extract_dispatch_numbers(
    rows: Iterable[Sequence[Any]],
    *,
    source_column: int = 1,
) -> tuple[str, ...]:
    """Read unique dispatch numbers from the downloaded first sheet's A column."""

    return extract_booking_numbers(rows, source_column=source_column, prefix="M")


def extract_truck_reservation_numbers(
    rows: Iterable[Sequence[Any]],
    *,
    source_column: int = 3,
) -> tuple[str, ...]:
    """Read unique Truck reservation keys from the downloaded C column."""

    return extract_booking_numbers(rows, source_column=source_column, prefix="T")


def _clean_text(value: Any) -> str:
    return re.sub(r"\s+", " ", str(value or "")).strip()


def _clean_vendor_name(value: Any) -> str:
    return _VENDOR_CODE_SUFFIX.sub("", _clean_text(value)).strip()


def _clean_order_number(value: Any) -> str:
    text = _clean_text(value)
    match = re.search(r"\d{5,20}", text)
    return match.group(0) if match else text


def parse_detail_table_cells(
    rows: Iterable[Sequence[Any]],
    *,
    dispatch_number: str = "",
    booking_prefix: str = "M",
) -> tuple[MilkrunProductRow, ...]:
    """Parse the booking detail tbody after its rowspans have been expanded.

    Milkrun details expose vendor/Milkrun/pallet/box plus the SKU columns.
    Truck details are read from ``#truckContainerList``. Rows for the same SKU
    are combined by both PALLET container count and total quantity so every SKU
    retains its own pallet and unit totals.
    """
    normalized_prefix = str(booking_prefix or "").strip().upper()
    if normalized_prefix == "T":
        aggregated: dict[str, dict[str, Any]] = {}
        order: list[str] = []
        for raw_cells in rows:
            cells = tuple(_clean_text(value) for value in raw_cells)
            if len(cells) < 9 or cells[0].upper() != "PALLET":
                continue
            sku_id = cells[4]
            sku_name = cells[5]
            if not sku_id and not sku_name:
                continue
            try:
                container_count = Decimal(cells[3].replace(",", ""))
                total_quantity = Decimal(cells[8].replace(",", ""))
            except InvalidOperation:
                continue
            if (
                not container_count.is_finite()
                or not total_quantity.is_finite()
                or container_count <= 0
                or total_quantity <= 0
            ):
                continue
            if sku_id not in aggregated:
                order.append(sku_id)
                aggregated[sku_id] = {
                    "sku_name": sku_name,
                    "container_count": Decimal("0"),
                    "total_quantity": Decimal("0"),
                    "container_names": [],
                }
            entry = aggregated[sku_id]
            entry["container_count"] += container_count
            entry["total_quantity"] += total_quantity
            if cells[1] and cells[1] not in entry["container_names"]:
                entry["container_names"].append(cells[1])

        def count_text(value: Decimal) -> str:
            text = format(value, "f")
            return text.rstrip("0").rstrip(".") if "." in text else text

        return tuple(
            MilkrunProductRow(
                vendor_name="",
                milkrun_number=", ".join(aggregated[sku_id]["container_names"]),
                pallet_count=count_text(aggregated[sku_id]["container_count"]),
                box_count=count_text(aggregated[sku_id]["total_quantity"]),
                sku_id=sku_id,
                sku_name=aggregated[sku_id]["sku_name"],
                dispatch_number=normalize_booking_number(
                    dispatch_number,
                    prefix="T",
                ),
            )
            for sku_id in order
        )

    if normalized_prefix != "M":
        raise ValueError(f"지원하지 않는 입고 예약 접두사입니다: {booking_prefix!r}")

    current_group: tuple[str, str, str] | None = None
    parsed: list[MilkrunProductRow] = []

    for raw_cells in rows:
        cells = tuple(_clean_text(value) for value in raw_cells)
        # The detail's group-level box count is intentionally ignored. WMS
        # ``hidden-weight`` is a one-unit weight, so each SKU row's final
        # confirmed-order quantity is the value used for pallet calculation.
        if len(cells) >= 9:
            current_group = (
                _clean_vendor_name(cells[0]),
                cells[1],
                cells[2],
            )
        if current_group is None or len(cells) < 5:
            continue

        if len(cells) >= 8:
            sku_id = cells[6]
            sku_name = cells[7]
            order_number = _clean_order_number(cells[4])
        else:
            sku_id = cells[-4]
            sku_name = cells[-3]
            order_number = _clean_order_number(cells[0])
        unit_count = cells[-1]
        if not sku_id and not sku_name:
            continue
        parsed.append(
            MilkrunProductRow(
                vendor_name=current_group[0],
                milkrun_number=current_group[1],
                pallet_count=current_group[2],
                # Field name retained for snapshot/JSON backward compatibility.
                box_count=unit_count,
                sku_id=sku_id,
                sku_name=sku_name,
                dispatch_number=normalize_booking_number(
                    dispatch_number,
                    prefix=booking_prefix,
                ),
                order_number=order_number,
            )
        )

    return tuple(parsed)
