from __future__ import annotations

import json
import tempfile
import unittest
from dataclasses import FrozenInstanceError
from decimal import Decimal
from pathlib import Path
from unittest.mock import patch

from Modules.WMS.ProductMemory import (
    GRAIN_CATEGORY,
    HEAVY_CATEGORY,
    HIGH_CATEGORY,
    LIGHT_CATEGORY,
    MEMORY_TYPE,
    MEMORY_VERSION,
    ProductMemory,
    calculate_boxes_per_pallet,
    calculate_pallet_measurement,
    normalize_product_name,
    normalize_sku_id,
    recover_manual_category_overrides,
)


class ProductMemoryTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temp_dir = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp_dir.cleanup)
        self.root = Path(self.temp_dir.name)

    def test_threshold_uses_unrounded_decimal_value(self) -> None:
        boxes, kilograms, category = calculate_pallet_measurement("1000", "279.999", "1")
        self.assertEqual(boxes, Decimal("279.999"))
        self.assertEqual(kilograms, Decimal("279.999"))
        self.assertEqual(category, LIGHT_CATEGORY)

        boxes, kilograms, category = calculate_pallet_measurement("1000", "280", "1")
        self.assertEqual(boxes, Decimal("280"))
        self.assertEqual(kilograms, Decimal("280"))
        self.assertEqual(category, HEAVY_CATEGORY)

    def test_boxes_per_pallet_is_decimal_and_does_not_require_wms_weight(self) -> None:
        self.assertEqual(calculate_boxes_per_pallet("2", "1"), Decimal("2"))
        self.assertEqual(calculate_boxes_per_pallet("1", "2"), Decimal("0.5"))

    def test_calculation_rejects_empty_zero_and_non_numeric_values(self) -> None:
        for values in (("", "10", "1"), ("1000", "0", "1"), ("1000", "10", "0"), ("x", "10", "1")):
            with self.subTest(values=values), self.assertRaises(ValueError):
                calculate_pallet_measurement(*values)

    def test_sku_normalization_handles_commas_dot_zero_and_rejects_non_ids(self) -> None:
        self.assertEqual(normalize_sku_id("72,246,083.0"), "72246083")
        self.assertEqual(normalize_sku_id(72246083.0), "72246083")
        self.assertEqual(normalize_sku_id(Decimal("72246083.0")), "72246083")
        self.assertEqual(normalize_sku_id("00123"), "00123")
        for value in (None, True, "", "0", "12.5", "SKU123"):
            with self.subTest(value=value), self.assertRaises(ValueError):
                normalize_sku_id(value)

    def test_product_name_collapses_line_break_without_losing_slash(self) -> None:
        self.assertEqual(normalize_product_name("  A/\r\n  B  "), "A/ B")
        self.assertEqual(normalize_product_name("A /\nB"), "A / B")

    def test_json_round_trip_preserves_slash_decimals_and_manual_override(self) -> None:
        memory = ProductMemory(self.root / "memory.json")
        automatic = memory.upsert_measurement("1,234.0", "상품 A/\n상품 B", "2000", "140", "1")
        overridden = memory.set_manual_category("1234", HIGH_CATEGORY)

        self.assertEqual(automatic.automatic_category, HEAVY_CATEGORY)
        self.assertIsNotNone(overridden)
        self.assertEqual(overridden.product_name, "상품 A/ 상품 B")
        self.assertEqual(overridden.effective_category, HIGH_CATEGORY)
        with self.assertRaises(FrozenInstanceError):
            overridden.product_name = "변경"

        exported = memory.export_to(self.root / "export.json")
        exported_text = exported.read_text(encoding="utf-8")
        self.assertIn("상품 A/ 상품 B", exported_text)
        self.assertNotIn("password", exported_text.lower())
        payload = json.loads(exported_text)
        self.assertEqual(payload["type"], MEMORY_TYPE)
        self.assertEqual(payload["version"], MEMORY_VERSION)

        restored = ProductMemory(exported)
        record = restored.get("1234")
        self.assertIsNotNone(record)
        self.assertEqual(record.weight_grams, Decimal("2000"))
        self.assertEqual(record.pallet_weight_kg, Decimal("280"))
        self.assertEqual(record.effective_category, HIGH_CATEGORY)

    def test_manual_categories_recover_when_unrelated_record_breaks_full_validation(self) -> None:
        path = self.root / "memory.json"
        memory = ProductMemory(path)
        memory.set_manual_category("101", HIGH_CATEGORY, "고단 상품")
        memory.set_manual_category("102", GRAIN_CATEGORY, "양곡 상품")
        payload = json.loads(path.read_text(encoding="utf-8"))
        payload["entries"].append(
            {
                "sku_id": "broken",
                "category_override": "지원하지 않는 분류",
            }
        )
        path.write_text(json.dumps(payload, ensure_ascii=False), encoding="utf-8")

        with self.assertRaises(ValueError):
            ProductMemory(path)

        recovered, skipped = recover_manual_category_overrides(path)

        self.assertEqual(recovered, {"101": HIGH_CATEGORY, "102": GRAIN_CATEGORY})
        self.assertEqual(skipped, 1)

    def test_manual_category_recovery_rejects_wrong_memory_identity(self) -> None:
        path = self.root / "memory.json"
        path.write_text(
            json.dumps(
                {
                    "type": "AnotherApplication",
                    "version": MEMORY_VERSION,
                    "entries": [
                        {"sku_id": "101", "category_override": HIGH_CATEGORY}
                    ],
                },
                ensure_ascii=False,
            ),
            encoding="utf-8",
        )

        with self.assertRaisesRegex(ValueError, "지원하지 않는 상품 메모리 형식"):
            recover_manual_category_overrides(path)

    def test_manual_category_can_create_unmeasured_placeholder(self) -> None:
        memory = ProductMemory(self.root / "memory.json")
        record = memory.set_manual_category("9,999", HIGH_CATEGORY, "A/\nB")

        self.assertIsNotNone(record)
        self.assertEqual(record.sku_id, "9999")
        self.assertEqual(record.product_name, "A/ B")
        self.assertIsNone(record.weight_grams)
        self.assertEqual(record.automatic_category, "")
        self.assertEqual(record.effective_category, HIGH_CATEGORY)

        payload = json.loads((self.root / "memory.json").read_text(encoding="utf-8"))
        self.assertIsNone(payload["entries"][0]["weight_grams"])
        self.assertIsNone(payload["entries"][0]["boxes_per_pallet"])
        restored = ProductMemory(self.root / "memory.json").get("9999")
        self.assertIsNotNone(restored)
        self.assertIsNone(restored.weight_grams)
        with self.assertRaises(ValueError):
            memory.update_calculation("9999", "10", "1")

    def test_clearing_override_deletes_an_unmeasured_placeholder(self) -> None:
        memory = ProductMemory(self.root / "memory.json")
        memory.set_manual_category("9999", LIGHT_CATEGORY, "상품")

        self.assertIsNone(memory.set_manual_category("9999", None))
        self.assertIsNone(memory.get("9999"))

    def test_upsert_measurement_keeps_manual_override_and_updates_calculation(self) -> None:
        memory = ProductMemory(self.root / "memory.json")
        memory.set_manual_category("123", HIGH_CATEGORY, "수동 상품")
        measured = memory.upsert_measurement("123", "측정 상품", "1000", "100", "1")

        self.assertEqual(measured.automatic_category, LIGHT_CATEGORY)
        self.assertEqual(measured.effective_category, HIGH_CATEGORY)
        self.assertEqual(measured.product_name, "측정 상품")

        recalculated = memory.update_calculation("123", "300", "1")
        self.assertEqual(recalculated.automatic_category, HEAVY_CATEGORY)
        self.assertEqual(recalculated.effective_category, HIGH_CATEGORY)
        cleared = memory.set_manual_category("123", None)
        self.assertIsNotNone(cleared)
        self.assertEqual(cleared.effective_category, HEAVY_CATEGORY)

    def test_cached_weight_recalculates_two_boxes_on_one_pallet_and_keeps_override(self) -> None:
        memory = ProductMemory(self.root / "memory.json")
        memory.upsert_measurement("123", "상품", "1000", "80", "2")
        memory.set_manual_category("123", HIGH_CATEGORY)

        recalculated = memory.update_calculation("123", "2", "1")

        self.assertEqual(recalculated.boxes_per_pallet, Decimal("2"))
        self.assertEqual(recalculated.pallet_weight_kg, Decimal("2"))
        self.assertEqual(recalculated.automatic_category, LIGHT_CATEGORY)
        self.assertEqual(recalculated.category_override, HIGH_CATEGORY)

    def test_grain_override_uses_the_same_persistent_process_as_high(self) -> None:
        memory = ProductMemory(self.root / "memory.json")
        memory.upsert_measurement("123", "양곡 상품", "1000", "300", "1")
        memory.set_manual_category("123", GRAIN_CATEGORY)

        recalculated = memory.update_calculation("123", "2", "1")
        refreshed = memory.upsert_weight("123", "양곡/ 상품", "1000")

        self.assertEqual(recalculated.automatic_category, LIGHT_CATEGORY)
        self.assertEqual(recalculated.category_override, GRAIN_CATEGORY)
        self.assertEqual(refreshed.category_override, GRAIN_CATEGORY)
        self.assertEqual(refreshed.effective_category, GRAIN_CATEGORY)

        exported = memory.export_to(self.root / "grain.json")
        restored = ProductMemory(exported).get("123")
        self.assertIsNotNone(restored)
        self.assertEqual(restored.category_override, GRAIN_CATEGORY)

    def test_cached_weight_recalculation_clears_light_or_heavy_manual_override(self) -> None:
        memory = ProductMemory(self.root / "memory.json")
        memory.upsert_measurement("123", "상품", "1000", "300", "1")
        memory.set_manual_category("123", HEAVY_CATEGORY)

        recalculated = memory.update_calculation("123", "2", "1")

        self.assertEqual(recalculated.boxes_per_pallet, Decimal("2"))
        self.assertEqual(recalculated.pallet_weight_kg, Decimal("2"))
        self.assertEqual(recalculated.automatic_category, LIGHT_CATEGORY)
        self.assertIsNone(recalculated.category_override)
        self.assertEqual(recalculated.effective_category, LIGHT_CATEGORY)

    def test_new_measurement_and_weight_refresh_keep_persistent_manual_override(self) -> None:
        memory = ProductMemory(self.root / "memory.json")
        memory.set_manual_category("123", LIGHT_CATEGORY, "상품")

        measured = memory.upsert_measurement("123", "상품", "1000", "300", "1")
        self.assertIsNone(measured.category_override)
        self.assertEqual(measured.effective_category, HEAVY_CATEGORY)

        memory.set_manual_category("123", HEAVY_CATEGORY)
        weight_only = memory.upsert_weight("123", "상품", "2000")
        self.assertIsNone(weight_only.category_override)

        memory.set_manual_category("123", HIGH_CATEGORY)
        high = memory.upsert_weight("123", "상품", "3000")
        self.assertEqual(high.category_override, HIGH_CATEGORY)

    def test_weight_only_update_preserves_existing_category_and_calculation(self) -> None:
        memory = ProductMemory(self.root / "memory.json")
        memory.upsert_measurement("123", "기존 상품", "1000", "300", "1")
        memory.set_manual_category("123", HIGH_CATEGORY)

        weight_only = memory.upsert_weight_only("123", "트럭 상품", "1000")

        self.assertEqual(weight_only.product_name, "트럭 상품")
        self.assertEqual(weight_only.weight_grams, Decimal("1000"))
        self.assertEqual(weight_only.automatic_category, HEAVY_CATEGORY)
        self.assertEqual(weight_only.category_override, HIGH_CATEGORY)
        self.assertEqual(weight_only.boxes_per_pallet, Decimal("300"))
        self.assertEqual(weight_only.pallet_weight_kg, Decimal("300"))
        self.assertEqual(weight_only.effective_category, HIGH_CATEGORY)
        self.assertEqual(ProductMemory(self.root / "memory.json").get("123"), weight_only)

    def test_weight_only_update_rejects_changed_weight_that_would_corrupt_calculation(self) -> None:
        path = self.root / "memory.json"
        memory = ProductMemory(path)
        original = memory.upsert_measurement("123", "상품", "1000", "300", "1")
        original_bytes = path.read_bytes()

        with self.assertRaisesRegex(ValueError, "계산값을 유지하면서 WMS 무게"):
            memory.upsert_weight_only("123", "상품", "2000")

        self.assertEqual(path.read_bytes(), original_bytes)
        self.assertEqual(memory.get("123"), original)
        self.assertEqual(ProductMemory(path).get("123"), original)

    def test_weight_only_update_adds_weight_without_deleting_manual_placeholder(self) -> None:
        memory = ProductMemory(self.root / "memory.json")
        placeholder = memory.set_manual_category("456", HIGH_CATEGORY, "미측정 상품")

        updated = memory.upsert_weight_only("456", "측정 상품", "1000")

        self.assertIsNotNone(placeholder)
        self.assertEqual(updated.product_name, "측정 상품")
        self.assertEqual(updated.weight_grams, Decimal("1000"))
        self.assertEqual(updated.automatic_category, "")
        self.assertEqual(updated.category_override, HIGH_CATEGORY)
        self.assertIsNone(updated.boxes_per_pallet)
        self.assertIsNone(updated.pallet_weight_kg)
        self.assertEqual(updated.effective_category, HIGH_CATEGORY)
        self.assertEqual(ProductMemory(self.root / "memory.json").get("456"), updated)

    def test_legacy_light_or_heavy_override_json_remains_readable_until_next_calculation(self) -> None:
        path = self.root / "memory.json"
        memory = ProductMemory(path)
        memory.upsert_measurement("123", "상품", "1000", "300", "1")
        memory.set_manual_category("123", LIGHT_CATEGORY)

        restored = ProductMemory(path).get("123")

        self.assertIsNotNone(restored)
        self.assertEqual(restored.category_override, LIGHT_CATEGORY)
        recalculated = ProductMemory(path).update_calculation("123", "300", "1")
        self.assertIsNone(recalculated.category_override)
        self.assertEqual(recalculated.effective_category, HEAVY_CATEGORY)

    def test_weight_only_cache_survives_invalid_pallet_inputs_and_can_recalculate_later(self) -> None:
        path = self.root / "memory.json"
        memory = ProductMemory(path)
        weight_only = memory.upsert_weight("456", "A/\nB", "2,000")

        self.assertEqual(weight_only.weight_grams, Decimal("2000"))
        self.assertEqual(weight_only.product_name, "A/ B")
        self.assertIsNone(weight_only.boxes_per_pallet)
        self.assertIsNone(weight_only.pallet_weight_kg)
        self.assertEqual(weight_only.automatic_category, "")
        self.assertEqual(weight_only.effective_category, "")

        restored = ProductMemory(path).get("456")
        self.assertIsNotNone(restored)
        self.assertEqual(restored.weight_grams, Decimal("2000"))
        with self.assertRaises(ValueError):
            memory.update_calculation("456", "0", "1")
        self.assertEqual(memory.get("456"), weight_only)

        calculated = memory.update_calculation("456", "140", "1")
        self.assertEqual(calculated.pallet_weight_kg, Decimal("280"))
        self.assertEqual(calculated.effective_category, HEAVY_CATEGORY)

    def test_atomic_replace_failure_preserves_file_and_memory(self) -> None:
        path = self.root / "memory.json"
        memory = ProductMemory(path)
        memory.upsert_measurement("100", "기존", "1000", "100", "1")
        original_bytes = path.read_bytes()

        with patch("Modules.WMS.ProductMemory.os.replace", side_effect=OSError("replace failed")):
            with self.assertRaisesRegex(OSError, "replace failed"):
                memory.upsert_measurement("200", "신규", "1000", "100", "1")

        self.assertEqual(path.read_bytes(), original_bytes)
        self.assertIsNone(memory.get("200"))
        self.assertEqual(ProductMemory(path).get("100").product_name, "기존")
        self.assertEqual(list(self.root.glob("*.tmp")), [])

    def test_corrupt_import_is_fully_rejected_without_mutation(self) -> None:
        path = self.root / "memory.json"
        memory = ProductMemory(path)
        memory.upsert_measurement("100", "기존", "1000", "100", "1")
        original_bytes = path.read_bytes()

        corrupt = self.root / "corrupt.json"
        corrupt.write_text('{"type": "UnHelper_milkrun_sku_memory", "version": 1, "entries": [', encoding="utf-8")
        with self.assertRaises(ValueError):
            memory.import_from(corrupt)

        self.assertEqual(path.read_bytes(), original_bytes)
        self.assertEqual([record.sku_id for record in memory.entries()], ["100"])

    def test_import_validates_every_entry_then_keeps_existing_sku(self) -> None:
        destination = ProductMemory(self.root / "destination.json")
        destination.upsert_measurement("100", "기존 우선", "1000", "100", "1")

        source = ProductMemory(self.root / "source.json")
        source.upsert_measurement("100", "가져온 중복", "2000", "140", "1")
        source.upsert_measurement("200", "신규", "1000", "50", "1")
        export_path = source.export_to(self.root / "import.json")

        summary = destination.import_from(export_path)
        self.assertEqual((summary.added, summary.skipped, summary.total), (1, 1, 2))
        self.assertEqual(destination.get("100").product_name, "기존 우선")
        self.assertEqual(destination.get("200").product_name, "신규")

        payload = json.loads(export_path.read_text(encoding="utf-8"))
        payload["entries"].append(
            {
                "sku_id": "300",
                "product_name": "잘못된 수동 레코드",
                "weight_grams": None,
                "automatic_category": "",
                "category_override": None,
                "boxes_per_pallet": None,
                "pallet_weight_kg": None,
                "measured_at": "",
                "updated_at": payload["exported_at"],
            }
        )
        invalid_path = self.root / "invalid-manual.json"
        invalid_path.write_text(json.dumps(payload, ensure_ascii=False), encoding="utf-8")
        before = destination.entries()
        with self.assertRaises(ValueError):
            destination.import_from(invalid_path)
        self.assertEqual(destination.entries(), before)
        self.assertIsNone(destination.get("300"))

    def test_import_records_can_overwrite_only_user_selected_duplicate_skus(self) -> None:
        destination = ProductMemory(self.root / "destination-overwrite.json")
        destination.upsert_measurement("100", "현재 100", "1000", "100", "1")
        destination.upsert_measurement("200", "현재 200", "900", "100", "1")

        source = ProductMemory(self.root / "source-overwrite.json")
        source.upsert_measurement("100", "가져온 100", "2000", "150", "1")
        source.upsert_measurement("200", "가져온 200", "3000", "100", "1")
        source.upsert_measurement("300", "신규 300", "500", "100", "1")

        summary = destination.import_records(
            source.entries(),
            overwrite_sku_ids={"100"},
        )

        self.assertEqual(
            (summary.added, summary.overwritten, summary.skipped, summary.total),
            (1, 1, 1, 3),
        )
        self.assertEqual(destination.get("100").product_name, "가져온 100")
        self.assertEqual(destination.get("100").weight_grams, Decimal("2000"))
        self.assertEqual(destination.get("200").product_name, "현재 200")
        self.assertEqual(destination.get("200").weight_grams, Decimal("900"))
        self.assertEqual(destination.get("300").product_name, "신규 300")

    def test_corrupt_cache_can_be_quarantined_and_reset_without_losing_backup(self) -> None:
        path = self.root / "memory.json"
        corrupt_bytes = b'{"type": "broken", "password": "must-stay-local"'
        path.write_bytes(corrupt_bytes)

        backup = ProductMemory.quarantine_and_reset(path)

        self.assertTrue(backup.is_file())
        self.assertEqual(backup.read_bytes(), corrupt_bytes)
        self.assertIn(".corrupt_", backup.name)
        self.assertEqual(ProductMemory(path).entries(), ())
        payload = json.loads(path.read_text(encoding="utf-8"))
        self.assertEqual(payload["type"], MEMORY_TYPE)
        self.assertEqual(payload["entries"], [])


if __name__ == "__main__":
    unittest.main()
