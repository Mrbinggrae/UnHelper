from __future__ import annotations

import json
import os
import tempfile
from dataclasses import dataclass, replace
from datetime import date, datetime
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import Any, Iterable, Mapping

from Modules.Shipments.DailyInbound import (
    MilkrunProductRow,
    normalize_truck_card_number,
)


STORE_TYPE = "UnHelper_raw_table_snapshots"
BUNDLE_TYPE = "UnHelper_raw_table_bundle"
FORMAT_VERSION = 2
SUPPORTED_FORMAT_VERSIONS = frozenset({1, FORMAT_VERSION})
MAX_SNAPSHOT_DATES = 2
MAX_ROWS_PER_TABLE = 5_000
MAX_FILE_BYTES = 10 * 1024 * 1024
MAX_CONTAINER_PALLETS = Decimal("1000000")
BOOKING_TYPES = ("milkrun", "truck")

_STORE_KEYS = frozenset({"type", "version", "entries"})
_BUNDLE_KEYS = frozenset(
    {"type", "version", "exported_at", "snapshot", "product_memory"}
)
_SNAPSHOT_KEYS = frozenset({"base_date", "updated_at", "tables"})
_BASE_PRODUCT_KEYS = frozenset(
    {
        "vendor_name",
        "milkrun_number",
        "pallet_count",
        "box_count",
        "sku_id",
        "sku_name",
        "dispatch_number",
    }
)
_PRODUCT_KEYS = _BASE_PRODUCT_KEYS | {"container_pallets", "detail_unavailable"}
_LEGACY_PRODUCT_KEYS = _BASE_PRODUCT_KEYS | {"order_number"}


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


def _validated_format_version(value: Any, file_label: str) -> int:
    if type(value) is not int or value not in SUPPORTED_FORMAT_VERSIONS:
        raise ValueError(
            f"지원하지 않는 {file_label} 버전입니다: {value!r}"
        )
    return value


def _product_to_json(product: MilkrunProductRow) -> dict[str, Any]:
    return {
        "vendor_name": str(product.vendor_name or ""),
        "milkrun_number": str(product.milkrun_number or ""),
        "pallet_count": str(product.pallet_count or ""),
        "box_count": str(product.box_count or ""),
        "sku_id": str(product.sku_id or ""),
        "sku_name": str(product.sku_name or ""),
        "dispatch_number": str(product.dispatch_number or ""),
        "container_pallets": [
            [str(identity), str(pallet_count)]
            for identity, pallet_count in product.container_pallets
        ],
        "detail_unavailable": product.detail_unavailable,
    }


def _product_from_json(
    raw: Any,
    location: str,
    *,
    format_version: int,
    booking_type: str,
) -> MilkrunProductRow:
    if not isinstance(raw, dict):
        raise ValueError(f"{location} 행이 객체 형식이 아닙니다.")
    raw_keys = set(raw)
    allowed_keys = (
        (_BASE_PRODUCT_KEYS, _LEGACY_PRODUCT_KEYS)
        if format_version == 1
        else (_PRODUCT_KEYS,)
    )
    if raw_keys not in allowed_keys:
        raise ValueError(f"{location} 행의 필드 구성이 올바르지 않습니다.")
    values: dict[str, str] = {}
    for key in _BASE_PRODUCT_KEYS:
        value = raw[key]
        if not isinstance(value, str):
            raise ValueError(f"{location}.{key} 값은 문자열이어야 합니다.")
        if len(value) > 10_000:
            raise ValueError(f"{location}.{key} 값이 너무 깁니다.")
        values[key] = value
    raw_container_pallets = raw.get("container_pallets", [])
    if not isinstance(raw_container_pallets, list) or len(raw_container_pallets) > 5_000:
        raise ValueError(f"{location}.container_pallets 구성이 올바르지 않습니다.")
    container_pallets: list[tuple[str, str]] = []
    seen_container_ids: set[str] = set()
    for index, allocation in enumerate(raw_container_pallets):
        if (
            not isinstance(allocation, list)
            or len(allocation) != 2
            or not all(isinstance(value, str) for value in allocation)
        ):
            raise ValueError(
                f"{location}.container_pallets[{index}] 구성이 올바르지 않습니다."
            )
        identity, pallet_count = (value.strip() for value in allocation)
        if not identity or len(identity) > 512 or len(pallet_count) > 100:
            raise ValueError(
                f"{location}.container_pallets[{index}] 값이 올바르지 않습니다."
            )
        try:
            parsed_count = Decimal(pallet_count.replace(",", ""))
        except InvalidOperation as exc:
            raise ValueError(
                f"{location}.container_pallets[{index}] 팔렛트 수가 숫자가 아닙니다."
            ) from exc
        if (
            not parsed_count.is_finite()
            or parsed_count <= 0
            or parsed_count > MAX_CONTAINER_PALLETS
            or parsed_count != parsed_count.to_integral_value()
        ):
            raise ValueError(
                f"{location}.container_pallets[{index}] 팔렛트 수는 "
                f"1~{MAX_CONTAINER_PALLETS} 범위의 정수여야 합니다."
            )
        normalized_identity = identity.casefold()
        if normalized_identity in seen_container_ids:
            raise ValueError(
                f"{location}.container_pallets에 중복 컨테이너가 있습니다: {identity}"
            )
        seen_container_ids.add(normalized_identity)
        container_pallets.append((identity, format(parsed_count.to_integral_value(), "f")))
    detail_unavailable = raw.get("detail_unavailable", False)
    if not isinstance(detail_unavailable, bool):
        raise ValueError(f"{location}.detail_unavailable 값은 참/거짓이어야 합니다.")
    if detail_unavailable:
        if (
            booking_type != "truck"
            or not normalize_truck_card_number(values["dispatch_number"])
        ):
            raise ValueError(
                f"{location}.detail_unavailable 행은 트럭 표에 있으며 "
                "T로 시작하는 트럭 예약번호가 필요합니다."
            )
        if values["sku_id"].strip():
            raise ValueError(
                f"{location}.detail_unavailable 행의 SKU ID는 비어 있어야 합니다."
            )
        if container_pallets:
            raise ValueError(
                f"{location}.detail_unavailable 행의 컨테이너 팔렛트는 "
                "비어 있어야 합니다."
            )
    return MilkrunProductRow(
        **values,
        container_pallets=tuple(container_pallets),
        detail_unavailable=detail_unavailable,
    )


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


def _snapshot_from_json(
    raw: Any,
    location: str,
    *,
    format_version: int,
) -> BookingDateSnapshot:
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
            _product_from_json(
                row,
                f"{location}.tables.{booking_type}[{index}]",
                format_version=format_version,
                booking_type=booking_type,
            )
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
        format_version = _validated_format_version(
            payload.get("version"),
            "RAW 표 저장",
        )
        entries = payload.get("entries")
        if not isinstance(entries, list):
            raise ValueError("RAW 표 저장 entries는 배열이어야 합니다.")
        if len(entries) > MAX_SNAPSHOT_DATES:
            raise ValueError("RAW 표 저장 파일에는 기준일을 최대 2개까지 저장할 수 있습니다.")
        result: dict[date, BookingDateSnapshot] = {}
        for index, raw in enumerate(entries):
            snapshot = _snapshot_from_json(
                raw,
                f"entries[{index}]",
                format_version=format_version,
            )
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
        validated = _snapshot_from_json(
            _snapshot_to_json(updated),
            "snapshot",
            format_version=FORMAT_VERSION,
        )
        snapshots[base_date] = validated
        self._commit(snapshots.values())
        return validated

    def save_snapshot(self, snapshot: BookingDateSnapshot) -> BookingDateSnapshot:
        validated = _snapshot_from_json(
            _snapshot_to_json(snapshot),
            "snapshot",
            format_version=FORMAT_VERSION,
        )
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
        format_version = _validated_format_version(
            payload.get("version"),
            "RAW 표 공유",
        )
        _validated_timestamp(payload["exported_at"], "exported_at")
        snapshot = _snapshot_from_json(
            payload["snapshot"],
            "snapshot",
            format_version=format_version,
        )
        product_memory = payload["product_memory"]
        if not isinstance(product_memory, dict):
            raise ValueError("RAW 표 공유 파일의 상품 메모리가 객체 형식이 아닙니다.")
        return snapshot, product_memory
