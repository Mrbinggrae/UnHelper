from __future__ import annotations

import json
import os
import tempfile
from dataclasses import dataclass, replace
from datetime import date, datetime
from pathlib import Path
from typing import Any, Iterable, Mapping

from Modules.Shipments.DailyInbound import MilkrunProductRow


STORE_TYPE = "UnHelper_raw_table_snapshots"
BUNDLE_TYPE = "UnHelper_raw_table_bundle"
FORMAT_VERSION = 1
MAX_SNAPSHOT_DATES = 2
MAX_ROWS_PER_TABLE = 5_000
MAX_FILE_BYTES = 10 * 1024 * 1024
BOOKING_TYPES = ("milkrun", "truck")

_STORE_KEYS = frozenset({"type", "version", "entries"})
_BUNDLE_KEYS = frozenset(
    {"type", "version", "exported_at", "snapshot", "product_memory"}
)
_SNAPSHOT_KEYS = frozenset({"base_date", "updated_at", "tables"})
_PRODUCT_KEYS = frozenset(
    {
        "vendor_name",
        "milkrun_number",
        "pallet_count",
        "box_count",
        "sku_id",
        "sku_name",
        "dispatch_number",
        "order_number",
    }
)


@dataclass(frozen=True, slots=True)
class BookingDateSnapshot:
    base_date: date
    updated_at: str
    milkrun_products: tuple[MilkrunProductRow, ...] = ()
    truck_products: tuple[MilkrunProductRow, ...] = ()

    def products_for(self, booking_type: str) -> tuple[MilkrunProductRow, ...]:
        _validate_booking_type(booking_type)
        return self.truck_products if booking_type == "truck" else self.milkrun_products


def _now_iso() -> str:
    return datetime.now().astimezone().isoformat(timespec="microseconds")


def _validate_booking_type(booking_type: str) -> None:
    if booking_type not in BOOKING_TYPES:
        raise ValueError(f"지원하지 않는 RAW 표 종류입니다: {booking_type!r}")


def _validated_timestamp(value: Any, field_name: str) -> str:
    if not isinstance(value, str) or not value.strip():
        raise ValueError(f"{field_name} 값이 비어 있거나 문자열이 아닙니다.")
    timestamp = value.strip()
    try:
        datetime.fromisoformat(timestamp.replace("Z", "+00:00"))
    except ValueError as exc:
        raise ValueError(f"{field_name} 값이 ISO 날짜 형식이 아닙니다.") from exc
    return timestamp


def _product_to_json(product: MilkrunProductRow) -> dict[str, str]:
    return {
        "vendor_name": str(product.vendor_name or ""),
        "milkrun_number": str(product.milkrun_number or ""),
        "pallet_count": str(product.pallet_count or ""),
        "box_count": str(product.box_count or ""),
        "sku_id": str(product.sku_id or ""),
        "sku_name": str(product.sku_name or ""),
        "dispatch_number": str(product.dispatch_number or ""),
        "order_number": str(product.order_number or ""),
    }


def _product_from_json(raw: Any, location: str) -> MilkrunProductRow:
    if not isinstance(raw, dict):
        raise ValueError(f"{location} 행이 객체 형식이 아닙니다.")
    unexpected = set(raw) - _PRODUCT_KEYS
    missing = _PRODUCT_KEYS - set(raw)
    if unexpected or missing:
        raise ValueError(f"{location} 행의 필드 구성이 올바르지 않습니다.")
    values: dict[str, str] = {}
    for key in _PRODUCT_KEYS:
        value = raw[key]
        if not isinstance(value, str):
            raise ValueError(f"{location}.{key} 값은 문자열이어야 합니다.")
        if len(value) > 10_000:
            raise ValueError(f"{location}.{key} 값이 너무 깁니다.")
        values[key] = value
    return MilkrunProductRow(**values)


def _snapshot_to_json(snapshot: BookingDateSnapshot) -> dict[str, Any]:
    return {
        "base_date": snapshot.base_date.isoformat(),
        "updated_at": snapshot.updated_at,
        "tables": {
            booking_type: [
                _product_to_json(product)
                for product in snapshot.products_for(booking_type)
            ]
            for booking_type in BOOKING_TYPES
        },
    }


def _snapshot_from_json(raw: Any, location: str) -> BookingDateSnapshot:
    if not isinstance(raw, dict) or set(raw) != _SNAPSHOT_KEYS:
        raise ValueError(f"{location} 스냅샷 구성이 올바르지 않습니다.")
    try:
        base_date = date.fromisoformat(str(raw["base_date"]))
    except ValueError as exc:
        raise ValueError(f"{location}.base_date가 ISO 날짜 형식이 아닙니다.") from exc
    updated_at = _validated_timestamp(raw["updated_at"], f"{location}.updated_at")
    tables = raw["tables"]
    if not isinstance(tables, dict) or set(tables) != set(BOOKING_TYPES):
        raise ValueError(f"{location}.tables 구성이 올바르지 않습니다.")
    parsed: dict[str, tuple[MilkrunProductRow, ...]] = {}
    for booking_type in BOOKING_TYPES:
        rows = tables[booking_type]
        if not isinstance(rows, list):
            raise ValueError(f"{location}.tables.{booking_type}는 배열이어야 합니다.")
        if len(rows) > MAX_ROWS_PER_TABLE:
            raise ValueError(
                f"{booking_type} 표는 최대 {MAX_ROWS_PER_TABLE:,}행까지 저장할 수 있습니다."
            )
        parsed[booking_type] = tuple(
            _product_from_json(row, f"{location}.tables.{booking_type}[{index}]")
            for index, row in enumerate(rows)
        )
    return BookingDateSnapshot(
        base_date=base_date,
        updated_at=updated_at,
        milkrun_products=parsed["milkrun"],
        truck_products=parsed["truck"],
    )


def _read_json(path: Path) -> dict[str, Any]:
    with path.open("rb") as stream:
        raw_bytes = stream.read(MAX_FILE_BYTES + 1)
    if len(raw_bytes) > MAX_FILE_BYTES:
        raise ValueError("RAW 표 저장 파일은 10MB 이하여야 합니다.")
    try:
        payload = json.loads(raw_bytes.decode("utf-8-sig"))
    except (UnicodeDecodeError, json.JSONDecodeError) as exc:
        raise ValueError("RAW 표 저장 파일은 올바른 UTF-8 JSON이어야 합니다.") from exc
    if not isinstance(payload, dict):
        raise ValueError("RAW 표 저장 파일의 최상위 값은 객체여야 합니다.")
    return payload


def _atomic_write(path: Path, payload: Mapping[str, Any]) -> None:
    encoded = (json.dumps(payload, ensure_ascii=False, indent=2) + "\n").encode("utf-8")
    if len(encoded) > MAX_FILE_BYTES:
        raise ValueError("RAW 표 저장 파일은 10MB 이하여야 합니다.")
    path.parent.mkdir(parents=True, exist_ok=True)
    descriptor, temp_name = tempfile.mkstemp(
        dir=str(path.parent),
        prefix=f".{path.name}.",
        suffix=".tmp",
    )
    temp_path = Path(temp_name)
    try:
        with os.fdopen(descriptor, "wb") as stream:
            stream.write(encoded)
            stream.flush()
            os.fsync(stream.fileno())
        os.replace(temp_path, path)
    except BaseException:
        try:
            temp_path.unlink(missing_ok=True)
        except OSError:
            pass
        raise


class BookingSnapshotStore:
    def __init__(self, path: str | os.PathLike[str]) -> None:
        self.path = Path(path)

    def _load(self) -> dict[date, BookingDateSnapshot]:
        if not self.path.exists():
            return {}
        payload = _read_json(self.path)
        if set(payload) != _STORE_KEYS or payload.get("type") != STORE_TYPE:
            raise ValueError("UnHelper RAW 표 저장 파일이 아닙니다.")
        if payload.get("version") != FORMAT_VERSION:
            raise ValueError(
                f"지원하지 않는 RAW 표 저장 버전입니다: {payload.get('version')!r}"
            )
        entries = payload.get("entries")
        if not isinstance(entries, list):
            raise ValueError("RAW 표 저장 entries는 배열이어야 합니다.")
        if len(entries) > MAX_SNAPSHOT_DATES:
            raise ValueError("RAW 표 저장 파일에는 기준일을 최대 2개까지 저장할 수 있습니다.")
        result: dict[date, BookingDateSnapshot] = {}
        for index, raw in enumerate(entries):
            snapshot = _snapshot_from_json(raw, f"entries[{index}]")
            if snapshot.base_date in result:
                raise ValueError("RAW 표 저장 파일에 중복 기준일이 있습니다.")
            result[snapshot.base_date] = snapshot
        return result

    def _commit(self, snapshots: Iterable[BookingDateSnapshot]) -> None:
        ordered = sorted(
            snapshots,
            key=lambda snapshot: snapshot.updated_at,
            reverse=True,
        )[:MAX_SNAPSHOT_DATES]
        _atomic_write(
            self.path,
            {
                "type": STORE_TYPE,
                "version": FORMAT_VERSION,
                "entries": [_snapshot_to_json(snapshot) for snapshot in ordered],
            },
        )

    def get(self, base_date: date) -> BookingDateSnapshot | None:
        return self._load().get(base_date)

    def dates(self) -> tuple[date, ...]:
        return tuple(
            snapshot.base_date
            for snapshot in sorted(
                self._load().values(),
                key=lambda snapshot: snapshot.updated_at,
                reverse=True,
            )
        )

    def save_table(
        self,
        base_date: date,
        booking_type: str,
        products: Iterable[MilkrunProductRow],
    ) -> BookingDateSnapshot:
        _validate_booking_type(booking_type)
        snapshots = self._load()
        current = snapshots.get(
            base_date,
            BookingDateSnapshot(base_date=base_date, updated_at=_now_iso()),
        )
        rows = tuple(products)
        if len(rows) > MAX_ROWS_PER_TABLE:
            raise ValueError(
                f"{booking_type} 표는 최대 {MAX_ROWS_PER_TABLE:,}행까지 저장할 수 있습니다."
            )
        updated = replace(
            current,
            updated_at=_now_iso(),
            **{f"{booking_type}_products": rows},
        )
        snapshots[base_date] = updated
        self._commit(snapshots.values())
        return updated

    def save_snapshot(self, snapshot: BookingDateSnapshot) -> BookingDateSnapshot:
        validated = _snapshot_from_json(_snapshot_to_json(snapshot), "snapshot")
        snapshots = self._load()
        updated = replace(validated, updated_at=_now_iso())
        snapshots[updated.base_date] = updated
        self._commit(snapshots.values())
        return updated

    def export_bundle(
        self,
        base_date: date,
        destination: str | os.PathLike[str],
        product_memory_payload: Mapping[str, Any],
    ) -> Path:
        snapshot = self.get(base_date)
        if snapshot is None:
            raise ValueError(f"{base_date:%Y-%m-%d}에 저장된 RAW 표가 없습니다.")
        path = Path(destination)
        _atomic_write(
            path,
            {
                "type": BUNDLE_TYPE,
                "version": FORMAT_VERSION,
                "exported_at": _now_iso(),
                "snapshot": _snapshot_to_json(snapshot),
                "product_memory": dict(product_memory_payload),
            },
        )
        return path

    @staticmethod
    def read_bundle(
        source: str | os.PathLike[str],
    ) -> tuple[BookingDateSnapshot, dict[str, Any]]:
        payload = _read_json(Path(source))
        if set(payload) != _BUNDLE_KEYS or payload.get("type") != BUNDLE_TYPE:
            raise ValueError("UnHelper RAW 표 공유 파일이 아닙니다.")
        if payload.get("version") != FORMAT_VERSION:
            raise ValueError(
                f"지원하지 않는 RAW 표 공유 버전입니다: {payload.get('version')!r}"
            )
        _validated_timestamp(payload["exported_at"], "exported_at")
        snapshot = _snapshot_from_json(payload["snapshot"], "snapshot")
        product_memory = payload["product_memory"]
        if not isinstance(product_memory, dict):
            raise ValueError("RAW 표 공유 파일의 상품 메모리가 객체 형식이 아닙니다.")
        return snapshot, product_memory
