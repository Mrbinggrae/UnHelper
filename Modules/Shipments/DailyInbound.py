from __future__ import annotations

import re
from dataclasses import dataclass
from typing import Any, Iterable, Sequence


_ORDER_NUMBER_PATTERN = re.compile(r"^T?(\d{5,20})$", re.IGNORECASE)
_VENDOR_CODE_SUFFIX = re.compile(r"\s*\(\s*A[0-9A-Z_-]+\s*\)\s*$", re.IGNORECASE)


@dataclass(frozen=True)
class MilkrunProductRow:
    vendor_name: str
    milkrun_number: str
    pallet_count: str
    box_count: str
    sku_id: str
    sku_name: str
    order_number: str = ""


def normalize_order_number(value: Any) -> str:
    """Return the numeric key used by both Excel P values and schedule cards."""
    if value is None or isinstance(value, bool):
        return ""
    if isinstance(value, int):
        candidate = str(value)
    elif isinstance(value, float):
        candidate = str(int(value)) if value.is_integer() else str(value)
    else:
        candidate = str(value)

    candidate = candidate.split("/", 1)[0].strip().upper()
    candidate = re.sub(r"[\s,]+", "", candidate)
    if candidate.endswith(".0") and candidate[:-2].isdigit():
        candidate = candidate[:-2]
    match = _ORDER_NUMBER_PATTERN.fullmatch(candidate)
    return match.group(1) if match else ""


def extract_order_numbers(
    rows: Iterable[Sequence[Any]],
    *,
    source_column: int = 14,
) -> tuple[str, ...]:
    """Read target column P from values pasted at C (the source's 14th column)."""
    if source_column < 1:
        raise ValueError("source_column must be 1 or greater")

    index = source_column - 1
    seen: set[str] = set()
    result: list[str] = []
    for row in rows:
        if len(row) <= index:
            continue
        order_number = normalize_order_number(row[index])
        if order_number and order_number not in seen:
            seen.add(order_number)
            result.append(order_number)
    return tuple(result)


def _clean_text(value: Any) -> str:
    return re.sub(r"\s+", " ", str(value or "")).strip()


def _clean_vendor_name(value: Any) -> str:
    return _VENDOR_CODE_SUFFIX.sub("", _clean_text(value)).strip()


def parse_detail_table_cells(
    rows: Iterable[Sequence[Any]],
    *,
    order_number: str = "",
) -> tuple[MilkrunProductRow, ...]:
    """Parse the detail tbody, including five-cell rows continued by rowspan."""
    current_group: tuple[str, str, str, str] | None = None
    parsed: list[MilkrunProductRow] = []

    for raw_cells in rows:
        cells = tuple(_clean_text(value) for value in raw_cells)
        # A new vendor/milkrun group contributes the first four cells. The
        # remaining five cells are shipment id, image, SKU id/name, barcode,
        # and quantity; continuation rows contain only the final five cells.
        if len(cells) >= 9:
            current_group = (
                _clean_vendor_name(cells[0]),
                cells[1],
                cells[2],
                cells[3],
            )
        if current_group is None or len(cells) < 5:
            continue

        if len(cells) >= 8:
            sku_id = cells[6]
            sku_name = cells[7]
        else:
            sku_id = cells[-4]
            sku_name = cells[-3]
        if not sku_id and not sku_name:
            continue
        parsed.append(
            MilkrunProductRow(
                vendor_name=current_group[0],
                milkrun_number=current_group[1],
                pallet_count=current_group[2],
                box_count=current_group[3],
                sku_id=sku_id,
                sku_name=sku_name,
                order_number=normalize_order_number(order_number),
            )
        )

    return tuple(parsed)
