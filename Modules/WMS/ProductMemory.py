from __future__ import annotations

import json
import math
import os
import re
import tempfile
import threading
from dataclasses import dataclass, replace
from datetime import datetime
from decimal import Decimal, InvalidOperation, localcontext
from pathlib import Path
from typing import Any, Iterable, Mapping


MEMORY_TYPE = "UnHelper_milkrun_sku_memory"
MEMORY_VERSION = 1
MAX_FILE_BYTES = 10 * 1024 * 1024
MAX_JSON_BYTES = MAX_FILE_BYTES
MAX_ENTRIES = 100_000

LIGHT_CATEGORY = "경량"
HEAVY_CATEGORY = "중량"
HIGH_CATEGORY = "고단"
GRAIN_CATEGORY = "양곡"
AUTOMATIC_CATEGORIES = frozenset({LIGHT_CATEGORY, HEAVY_CATEGORY})
PERSISTENT_MANUAL_CATEGORIES = frozenset({HIGH_CATEGORY, GRAIN_CATEGORY})
MANUAL_CATEGORIES = frozenset(
    {LIGHT_CATEGORY, HEAVY_CATEGORY, HIGH_CATEGORY, GRAIN_CATEGORY}
)
HEAVY_THRESHOLD_KG = Decimal("280")
CALCULATION_PRECISION = 50

_NUMERIC_SKU = re.compile(r"\d+")
_WHOLE_DECIMAL_SUFFIX = re.compile(r"^(\d+)\.0+$")
_ROOT_KEYS = frozenset({"type", "version", "exported_at", "entries"})
_RECORD_KEYS = frozenset(
    {
        "sku_id",
        "product_name",
        "weight_grams",
        "automatic_category",
        "category_override",
        "boxes_per_pallet",
        "pallet_weight_kg",
        "measured_at",
        "updated_at",
    }
)


@dataclass(frozen=True, slots=True)
class ProductMemoryRecord:
    sku_id: str
    product_name: str
    weight_grams: Decimal | None
    automatic_category: str
    category_override: str | None
    boxes_per_pallet: Decimal | None
    pallet_weight_kg: Decimal | None
    measured_at: str
    updated_at: str

    @property
    def effective_category(self) -> str:
        return self.category_override or self.automatic_category


@dataclass(frozen=True, slots=True)
class ImportSummary:
    added: int
    skipped: int

    @property
    def total(self) -> int:
        return self.added + self.skipped


def normalize_sku_id(value: Any) -> str:
    """Return a positive, digits-only SKU key without Excel formatting."""
    if value is None or isinstance(value, bool):
        raise ValueError("SKU ID가 비어 있거나 올바르지 않습니다.")

    if isinstance(value, int):
        candidate = str(value)
    elif isinstance(value, float):
        if not math.isfinite(value) or not value.is_integer():
            raise ValueError(f"SKU ID는 숫자 정수여야 합니다: {value!r}")
        candidate = str(int(value))
    elif isinstance(value, Decimal):
        if not value.is_finite() or value != value.to_integral_value():
            raise ValueError(f"SKU ID는 숫자 정수여야 합니다: {value!r}")
        candidate = format(value, "f")
    else:
        candidate = str(value)

    candidate = re.sub(r"\s+", "", candidate.strip()).replace(",", "")
    whole_decimal = _WHOLE_DECIMAL_SUFFIX.fullmatch(candidate)
    if whole_decimal:
        candidate = whole_decimal.group(1)

    if not _NUMERIC_SKU.fullmatch(candidate) or int(candidate) <= 0:
        raise ValueError(f"SKU ID는 0보다 큰 숫자여야 합니다: {value!r}")
    return candidate


def normalize_product_name(value: Any) -> str:
    """Collapse visual line breaks while preserving slash characters."""
    return " ".join(str(value or "").split())


def _positive_decimal(value: Any, field_name: str) -> Decimal:
    if value is None or isinstance(value, bool):
        raise ValueError(f"{field_name} 값이 비어 있습니다.")
    text = str(value).strip().replace(",", "")
    if not text:
        raise ValueError(f"{field_name} 값이 비어 있습니다.")
    try:
        result = Decimal(text)
    except (InvalidOperation, ValueError) as exc:
        raise ValueError(f"{field_name} 값이 숫자가 아닙니다: {value!r}") from exc
    if not result.is_finite() or result <= 0:
        raise ValueError(f"{field_name} 값은 0보다 커야 합니다: {value!r}")
    return result


def calculate_boxes_per_pallet(
    box_count: Any,
    pallet_count: Any,
) -> Decimal:
    """Calculate SKU units per pallet without requiring WMS data.

    Legacy public/model names are retained for persisted JSON compatibility.
    """
    boxes = _positive_decimal(box_count, "유닛 수")
    pallets = _positive_decimal(pallet_count, "팔레트 수")

    with localcontext() as context:
        context.prec = CALCULATION_PRECISION
        return boxes / pallets


def calculate_pallet_measurement(
    weight_grams: Any,
    box_count: Any,
    pallet_count: Any,
) -> tuple[Decimal, Decimal, str]:
    """Calculate units/pallet, kg/pallet, and the unrounded threshold result."""
    weight = _positive_decimal(weight_grams, "상품 무게(g)")
    boxes_per_pallet = calculate_boxes_per_pallet(box_count, pallet_count)

    with localcontext() as context:
        context.prec = CALCULATION_PRECISION
        pallet_weight_kg = weight * boxes_per_pallet / Decimal("1000")
    automatic_category = HEAVY_CATEGORY if pallet_weight_kg >= HEAVY_THRESHOLD_KG else LIGHT_CATEGORY
    return boxes_per_pallet, pallet_weight_kg, automatic_category


def _now_iso() -> str:
    return datetime.now().astimezone().isoformat(timespec="seconds")


def _validated_timestamp(value: Any, field_name: str) -> str:
    if not isinstance(value, str) or not value.strip():
        raise ValueError(f"{field_name} 값이 비어 있거나 문자열이 아닙니다.")
    timestamp = value.strip()
    try:
        datetime.fromisoformat(timestamp.replace("Z", "+00:00"))
    except ValueError as exc:
        raise ValueError(f"{field_name} 값이 ISO 날짜 형식이 아닙니다: {value!r}") from exc
    return timestamp


def _validated_automatic_category(value: Any) -> str:
    if not isinstance(value, str) or value not in AUTOMATIC_CATEGORIES:
        raise ValueError("자동 분류는 경량 또는 중량이어야 합니다.")
    return value


def _validated_override(value: Any) -> str | None:
    if value is None:
        return None
    if not isinstance(value, str) or value not in MANUAL_CATEGORIES:
        raise ValueError("수동 분류는 경량, 중량, 고단, 양곡 또는 None이어야 합니다.")
    return value


def _override_after_automatic_recalculation(
    previous: ProductMemoryRecord | None,
) -> str | None:
    """Keep only the user classification that must survive new pallet inputs."""
    if (
        previous is not None
        and previous.category_override in PERSISTENT_MANUAL_CATEGORIES
    ):
        return previous.category_override
    return None


def _record_to_json(record: ProductMemoryRecord) -> dict[str, Any]:
    return {
        "sku_id": record.sku_id,
        "product_name": record.product_name,
        "weight_grams": str(record.weight_grams) if record.weight_grams is not None else None,
        "automatic_category": record.automatic_category,
        "category_override": record.category_override,
        "boxes_per_pallet": str(record.boxes_per_pallet) if record.boxes_per_pallet is not None else None,
        "pallet_weight_kg": str(record.pallet_weight_kg) if record.pallet_weight_kg is not None else None,
        "measured_at": record.measured_at,
        "updated_at": record.updated_at,
    }


def _record_from_json(raw: Any, index: int) -> ProductMemoryRecord:
    if not isinstance(raw, dict):
        raise ValueError(f"entries[{index}]가 객체 형식이 아닙니다.")
    unexpected = set(raw) - _RECORD_KEYS
    missing = _RECORD_KEYS - set(raw)
    if unexpected:
        raise ValueError(f"entries[{index}]에 허용되지 않은 키가 있습니다: {', '.join(sorted(unexpected))}")
    if missing:
        raise ValueError(f"entries[{index}]에 필수 키가 없습니다: {', '.join(sorted(missing))}")

    sku_id = normalize_sku_id(raw["sku_id"])
    if not isinstance(raw["product_name"], str):
        raise ValueError(f"entries[{index}].product_name은 문자열이어야 합니다.")
    product_name = normalize_product_name(raw["product_name"])
    category_override = _validated_override(raw["category_override"])
    updated_at = _validated_timestamp(raw["updated_at"], "updated_at")

    raw_weight = raw["weight_grams"]
    is_manual_placeholder = raw_weight is None or (isinstance(raw_weight, str) and not raw_weight.strip())
    if is_manual_placeholder:
        if category_override not in MANUAL_CATEGORIES:
            raise ValueError(f"entries[{index}]의 무게 없는 레코드에는 수동 분류가 필요합니다.")
        if raw["automatic_category"] not in (None, ""):
            raise ValueError(f"entries[{index}]의 무게 없는 레코드에는 자동 분류를 저장할 수 없습니다.")
        if raw["boxes_per_pallet"] not in (None, "") or raw["pallet_weight_kg"] not in (None, ""):
            raise ValueError(f"entries[{index}]의 무게 없는 레코드에는 계산값을 저장할 수 없습니다.")
        if raw["measured_at"] not in (None, ""):
            raise ValueError(f"entries[{index}]의 무게 없는 레코드에는 측정 시각을 저장할 수 없습니다.")
        weight_grams = None
        boxes_per_pallet = None
        pallet_weight_kg = None
        automatic_category = ""
        measured_at = ""
    else:
        weight_grams = _positive_decimal(raw_weight, "상품 무게(g)")
        measured_at = _validated_timestamp(raw["measured_at"], "measured_at")
        has_boxes = raw["boxes_per_pallet"] not in (None, "")
        has_pallet_weight = raw["pallet_weight_kg"] not in (None, "")
        if has_boxes != has_pallet_weight:
            raise ValueError(f"entries[{index}]의 팔레트 계산값이 일부만 저장되어 있습니다.")
        if not has_boxes:
            if raw["automatic_category"] not in (None, ""):
                raise ValueError(f"entries[{index}]의 계산값 없는 레코드에는 자동 분류를 저장할 수 없습니다.")
            boxes_per_pallet = None
            pallet_weight_kg = None
            automatic_category = ""
        else:
            boxes_per_pallet = _positive_decimal(raw["boxes_per_pallet"], "팔렛트당 유닛 수")
            pallet_weight_kg = _positive_decimal(raw["pallet_weight_kg"], "팔레트 무게(kg)")
            automatic_category = _validated_automatic_category(raw["automatic_category"])

            with localcontext() as context:
                context.prec = CALCULATION_PRECISION
                expected_weight = weight_grams * boxes_per_pallet / Decimal("1000")
            if pallet_weight_kg != expected_weight:
                raise ValueError(f"entries[{index}]의 팔레트 무게 계산값이 일치하지 않습니다.")
            expected_category = HEAVY_CATEGORY if pallet_weight_kg >= HEAVY_THRESHOLD_KG else LIGHT_CATEGORY
            if automatic_category != expected_category:
                raise ValueError(f"entries[{index}]의 자동 분류가 팔레트 무게와 일치하지 않습니다.")

    return ProductMemoryRecord(
        sku_id=sku_id,
        product_name=product_name,
        weight_grams=weight_grams,
        automatic_category=automatic_category,
        category_override=category_override,
        boxes_per_pallet=boxes_per_pallet,
        pallet_weight_kg=pallet_weight_kg,
        measured_at=measured_at,
        updated_at=updated_at,
    )


def _read_json_payload(path: Path) -> dict[str, Any]:
    with path.open("rb") as stream:
        raw_bytes = stream.read(MAX_FILE_BYTES + 1)
    if len(raw_bytes) > MAX_FILE_BYTES:
        raise ValueError(f"상품 메모리 파일은 {MAX_FILE_BYTES // (1024 * 1024)}MB 이하여야 합니다.")
    try:
        text = raw_bytes.decode("utf-8-sig")
    except UnicodeDecodeError as exc:
        raise ValueError("상품 메모리 파일은 UTF-8 JSON이어야 합니다.") from exc
    try:
        payload = json.loads(text)
    except json.JSONDecodeError as exc:
        raise ValueError(f"상품 메모리 JSON이 손상되었습니다: {exc}") from exc
    if not isinstance(payload, dict):
        raise ValueError("상품 메모리 JSON의 최상위 값은 객체여야 합니다.")
    return payload


def _validated_payload(payload: Mapping[str, Any]) -> tuple[list[ProductMemoryRecord], int]:
    unexpected = set(payload) - _ROOT_KEYS
    if unexpected:
        raise ValueError(f"상품 메모리 파일에 허용되지 않은 키가 있습니다: {', '.join(sorted(unexpected))}")
    if payload.get("type") != MEMORY_TYPE:
        raise ValueError("UnHelper 밀크런 상품 메모리 파일이 아닙니다.")
    if payload.get("version") != MEMORY_VERSION:
        raise ValueError(f"지원하지 않는 상품 메모리 버전입니다: {payload.get('version')!r}")
    if "exported_at" in payload:
        _validated_timestamp(payload["exported_at"], "exported_at")
    raw_entries = payload.get("entries")
    if not isinstance(raw_entries, list):
        raise ValueError("상품 메모리 entries는 배열이어야 합니다.")
    if len(raw_entries) > MAX_ENTRIES:
        raise ValueError(f"상품 메모리는 최대 {MAX_ENTRIES:,}개 항목까지 허용합니다.")

    records: list[ProductMemoryRecord] = []
    seen: set[str] = set()
    duplicate_count = 0
    for index, raw_record in enumerate(raw_entries):
        record = _record_from_json(raw_record, index)
        if record.sku_id in seen:
            duplicate_count += 1
            continue
        seen.add(record.sku_id)
        records.append(record)
    return records, duplicate_count


def _payload_for(records: Iterable[ProductMemoryRecord]) -> dict[str, Any]:
    return {
        "type": MEMORY_TYPE,
        "version": MEMORY_VERSION,
        "exported_at": _now_iso(),
        "entries": [_record_to_json(record) for record in records],
    }


def _atomic_write_json(path: Path, payload: Mapping[str, Any]) -> None:
    encoded = (json.dumps(payload, ensure_ascii=False, indent=2) + "\n").encode("utf-8")
    if len(encoded) > MAX_FILE_BYTES:
        raise ValueError(f"상품 메모리 파일은 {MAX_FILE_BYTES // (1024 * 1024)}MB 이하여야 합니다.")

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


def _record_sort_key(record: ProductMemoryRecord) -> tuple[int, str]:
    return len(record.sku_id), record.sku_id


class ProductMemory:
    def __init__(self, path: str | os.PathLike[str]) -> None:
        self.path = Path(path)
        self._lock = threading.RLock()
        self._records: dict[str, ProductMemoryRecord] = {}
        if self.path.exists():
            payload = _read_json_payload(self.path)
            records, duplicate_count = _validated_payload(payload)
            if duplicate_count:
                raise ValueError("저장된 상품 메모리에 중복 SKU ID가 있습니다.")
            self._records = {record.sku_id: record for record in records}

    def get(self, sku_id: Any) -> ProductMemoryRecord | None:
        key = normalize_sku_id(sku_id)
        with self._lock:
            return self._records.get(key)

    def entries(self) -> tuple[ProductMemoryRecord, ...]:
        with self._lock:
            return tuple(sorted(self._records.values(), key=_record_sort_key))

    def _commit(self, records: dict[str, ProductMemoryRecord]) -> None:
        ordered = sorted(records.values(), key=_record_sort_key)
        _atomic_write_json(self.path, _payload_for(ordered))
        self._records = records

    def upsert_measurement(
        self,
        sku_id: Any,
        product_name: Any,
        weight_grams: Any,
        box_count: Any,
        pallet_count: Any,
        *,
        measured_at: str | None = None,
    ) -> ProductMemoryRecord:
        key = normalize_sku_id(sku_id)
        name = normalize_product_name(product_name)
        weight = _positive_decimal(weight_grams, "상품 무게(g)")
        boxes_per_pallet, pallet_weight_kg, automatic_category = calculate_pallet_measurement(
            weight,
            box_count,
            pallet_count,
        )
        measured_timestamp = _validated_timestamp(measured_at, "measured_at") if measured_at else _now_iso()
        updated_timestamp = _now_iso()

        with self._lock:
            previous = self._records.get(key)
            record = ProductMemoryRecord(
                sku_id=key,
                product_name=name or (previous.product_name if previous else ""),
                weight_grams=weight,
                automatic_category=automatic_category,
                category_override=_override_after_automatic_recalculation(previous),
                boxes_per_pallet=boxes_per_pallet,
                pallet_weight_kg=pallet_weight_kg,
                measured_at=measured_timestamp,
                updated_at=updated_timestamp,
            )
            new_records = dict(self._records)
            new_records[key] = record
            self._commit(new_records)
            return record

    def upsert_weight(
        self,
        sku_id: Any,
        product_name: Any,
        weight_grams: Any,
        *,
        measured_at: str | None = None,
    ) -> ProductMemoryRecord:
        """Persist a successful WMS measurement before pallet inputs are usable."""
        return self._upsert_weight(
            sku_id,
            product_name,
            weight_grams,
            measured_at=measured_at,
            preserve_existing_state=False,
        )

    def upsert_weight_only(
        self,
        sku_id: Any,
        product_name: Any,
        weight_grams: Any,
        *,
        measured_at: str | None = None,
    ) -> ProductMemoryRecord:
        """Update WMS weight without changing existing classification/calculation state.

        A multi-SKU Truck reservation cannot safely attach a reservation-level
        unit/pallet calculation to each individual SKU. Existing global SKU
        classifications therefore remain untouched; a brand-new SKU is stored
        as a weight-only record.
        """
        return self._upsert_weight(
            sku_id,
            product_name,
            weight_grams,
            measured_at=measured_at,
            preserve_existing_state=True,
        )

    def _upsert_weight(
        self,
        sku_id: Any,
        product_name: Any,
        weight_grams: Any,
        *,
        measured_at: str | None,
        preserve_existing_state: bool,
    ) -> ProductMemoryRecord:
        key = normalize_sku_id(sku_id)
        name = normalize_product_name(product_name)
        weight = _positive_decimal(weight_grams, "상품 무게(g)")
        measured_timestamp = _validated_timestamp(measured_at, "measured_at") if measured_at else _now_iso()

        with self._lock:
            previous = self._records.get(key)
            if (
                preserve_existing_state
                and previous is not None
                and previous.boxes_per_pallet is not None
                and previous.weight_grams != weight
            ):
                raise ValueError(
                    f"SKU {key}의 기존 팔렛트 계산값을 유지하면서 WMS 무게를 "
                    "변경할 수 없습니다. 계산값을 먼저 다시 산출해야 합니다."
                )
            record = ProductMemoryRecord(
                sku_id=key,
                product_name=name or (previous.product_name if previous else ""),
                weight_grams=weight,
                automatic_category=(
                    previous.automatic_category
                    if preserve_existing_state and previous is not None
                    else ""
                ),
                category_override=(
                    previous.category_override
                    if preserve_existing_state and previous is not None
                    else _override_after_automatic_recalculation(previous)
                ),
                boxes_per_pallet=(
                    previous.boxes_per_pallet
                    if preserve_existing_state and previous is not None
                    else None
                ),
                pallet_weight_kg=(
                    previous.pallet_weight_kg
                    if preserve_existing_state and previous is not None
                    else None
                ),
                measured_at=measured_timestamp,
                updated_at=_now_iso(),
            )
            new_records = dict(self._records)
            new_records[key] = record
            self._commit(new_records)
            return record

    def update_calculation(
        self,
        sku_id: Any,
        box_count: Any,
        pallet_count: Any,
    ) -> ProductMemoryRecord:
        key = normalize_sku_id(sku_id)
        with self._lock:
            previous = self._records.get(key)
            if previous is None:
                raise KeyError(f"저장된 SKU ID가 아닙니다: {key}")
            if previous.weight_grams is None:
                raise ValueError(f"SKU {key}의 WMS 무게가 없어 계산할 수 없습니다.")
            boxes_per_pallet, pallet_weight_kg, automatic_category = calculate_pallet_measurement(
                previous.weight_grams,
                box_count,
                pallet_count,
            )
            record = replace(
                previous,
                boxes_per_pallet=boxes_per_pallet,
                pallet_weight_kg=pallet_weight_kg,
                automatic_category=automatic_category,
                category_override=_override_after_automatic_recalculation(previous),
                updated_at=_now_iso(),
            )
            new_records = dict(self._records)
            new_records[key] = record
            self._commit(new_records)
            return record

    def set_manual_category(
        self,
        sku_id: Any,
        category: str | None,
        product_name: Any = "",
    ) -> ProductMemoryRecord | None:
        key = normalize_sku_id(sku_id)
        category_override = _validated_override(category)
        normalized_name = normalize_product_name(product_name)
        with self._lock:
            previous = self._records.get(key)
            if previous is None:
                if category_override is None:
                    raise KeyError(f"저장된 SKU ID가 아닙니다: {key}")
                timestamp = _now_iso()
                record = ProductMemoryRecord(
                    sku_id=key,
                    product_name=normalized_name,
                    weight_grams=None,
                    automatic_category="",
                    category_override=category_override,
                    boxes_per_pallet=None,
                    pallet_weight_kg=None,
                    measured_at="",
                    updated_at=timestamp,
                )
            elif previous.weight_grams is None and category_override is None:
                new_records = dict(self._records)
                del new_records[key]
                self._commit(new_records)
                return None
            else:
                record = replace(
                    previous,
                    product_name=normalized_name or previous.product_name,
                    category_override=category_override,
                    updated_at=_now_iso(),
                )
            new_records = dict(self._records)
            new_records[key] = record
            self._commit(new_records)
            return record

    def delete(self, sku_id: Any) -> bool:
        key = normalize_sku_id(sku_id)
        with self._lock:
            if key not in self._records:
                return False
            new_records = dict(self._records)
            del new_records[key]
            self._commit(new_records)
            return True

    def export_to(self, path: str | os.PathLike[str]) -> Path:
        destination = Path(path)
        with self._lock:
            ordered = sorted(self._records.values(), key=_record_sort_key)
            _atomic_write_json(destination, _payload_for(ordered))
        return destination

    def export_payload(self, sku_ids: Iterable[Any] | None = None) -> dict[str, Any]:
        """Return a validated JSON payload, optionally limited to selected SKUs."""
        with self._lock:
            if sku_ids is None:
                records = tuple(self._records.values())
            else:
                selected: set[str] = set()
                for value in sku_ids:
                    try:
                        selected.add(normalize_sku_id(value))
                    except ValueError:
                        continue
                records = tuple(
                    record
                    for sku_id, record in self._records.items()
                    if sku_id in selected
                )
            return _payload_for(sorted(records, key=_record_sort_key))

    @staticmethod
    def validate_payload(payload: Mapping[str, Any]) -> tuple[ProductMemoryRecord, ...]:
        records, duplicate_count = _validated_payload(payload)
        if duplicate_count:
            raise ValueError("가져올 상품 메모리에 중복 SKU ID가 있습니다.")
        return tuple(records)

    def import_records(
        self,
        records: Iterable[ProductMemoryRecord],
    ) -> ImportSummary:
        validated = tuple(
            _record_from_json(_record_to_json(record), index)
            for index, record in enumerate(records)
        )
        with self._lock:
            new_records = dict(self._records)
            added = 0
            skipped = 0
            for record in validated:
                if record.sku_id in new_records:
                    skipped += 1
                    continue
                new_records[record.sku_id] = record
                added += 1
            if added:
                self._commit(new_records)
            return ImportSummary(added=added, skipped=skipped)

    def import_payload(self, payload: Mapping[str, Any]) -> ImportSummary:
        return self.import_records(self.validate_payload(payload))

    def import_from(self, path: str | os.PathLike[str]) -> ImportSummary:
        source = Path(path)
        payload = _read_json_payload(source)
        imported_records, duplicate_count = _validated_payload(payload)

        with self._lock:
            new_records = dict(self._records)
            added = 0
            skipped = duplicate_count
            for record in imported_records:
                if record.sku_id in new_records:
                    skipped += 1
                    continue
                new_records[record.sku_id] = record
                added += 1
            if added:
                self._commit(new_records)
            return ImportSummary(added=added, skipped=skipped)

    @classmethod
    def quarantine_and_reset(cls, path: str | os.PathLike[str]) -> Path:
        """Move an unreadable cache aside and replace it with a valid empty cache."""
        source = Path(path)
        if not source.is_file():
            raise FileNotFoundError(f"격리할 상품 메모리 파일을 찾을 수 없습니다: {source}")

        stamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
        suffix = source.suffix or ".json"
        backup = source.with_name(f"{source.stem}.corrupt_{stamp}{suffix}")
        sequence = 1
        while backup.exists():
            backup = source.with_name(f"{source.stem}.corrupt_{stamp}_{sequence}{suffix}")
            sequence += 1

        os.replace(source, backup)
        empty_memory = cls(source)
        with empty_memory._lock:
            empty_memory._commit({})
        return backup
