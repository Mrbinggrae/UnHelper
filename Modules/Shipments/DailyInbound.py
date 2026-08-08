from __future__ import annotations

import re
from dataclasses import dataclass
from typing import Any, Iterable, Sequence


_DISPATCH_NUMBER_PATTERN = re.compile(r"^M?(\d{5,20})$", re.IGNORECASE)
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


def normalize_dispatch_number(value: Any) -> str:
    """Return the canonical ``M<digits>`` key for Milkrun schedule cards."""
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
    if candidate.endswith(".0") and re.fullmatch(r"M?\d+", candidate[:-2]):
        candidate = candidate[:-2]
    match = _DISPATCH_NUMBER_PATTERN.fullmatch(candidate)
    return f"M{match.group(1)}" if match else ""


def normalize_milkrun_card_number(value: Any) -> str:
    """Accept only a schedule-card label that explicitly starts with ``M``."""

    if not isinstance(value, str):
        return ""
    candidate = re.sub(r"[\s,]+", "", value.strip().upper())
    if not candidate.startswith("M"):
        return ""
    return normalize_dispatch_number(candidate)


def extract_dispatch_numbers(
    rows: Iterable[Sequence[Any]],
    *,
    source_column: int = 1,
) -> tuple[str, ...]:
    """Read unique dispatch numbers from the downloaded first sheet's A column."""
    if source_column < 1:
        raise ValueError("source_column must be 1 or greater")

    index = source_column - 1
    seen: set[str] = set()
    result: list[str] = []
    for row in rows:
        if len(row) <= index:
            continue
        dispatch_number = normalize_dispatch_number(row[index])
        if dispatch_number and dispatch_number not in seen:
            seen.add(dispatch_number)
            result.append(dispatch_number)
    return tuple(result)


def _clean_text(value: Any) -> str:
    return re.sub(r"\s+", " ", str(value or "")).strip()


def _clean_vendor_name(value: Any) -> str:
    return _VENDOR_CODE_SUFFIX.sub("", _clean_text(value)).strip()


def parse_detail_table_cells(
    rows: Iterable[Sequence[Any]],
    *,
    dispatch_number: str = "",
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
                dispatch_number=normalize_dispatch_number(dispatch_number),
            )
        )

    return tuple(parsed)
