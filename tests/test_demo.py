import csv
import hashlib
import json
import tempfile
import unittest
from pathlib import Path

from price_demo import import_fingerprint, parse_number, run
from xlsx_io import read_xlsx_rows


ROOT = Path(__file__).resolve().parents[1]


class PriceDemoTests(unittest.TestCase):
    def test_parse_number_accepts_comma(self) -> None:
        self.assertEqual(parse_number("0,4"), 0.4)

    def test_mapping_version_is_part_of_import_fingerprint(self) -> None:
        supplier_mapping = {
            "supplier": "Supplier A",
            "columns": {"sku": "sku"},
        }
        source = ROOT / "samples" / "supplier_a.csv"
        self.assertNotEqual(
            import_fingerprint(source, supplier_mapping, "1.0.0"),
            import_fingerprint(source, supplier_mapping, "2.0.0"),
        )

    def test_xlsx_sample_is_real_readable_input(self) -> None:
        rows = read_xlsx_rows(ROOT / "samples" / "supplier_b.xlsx", "Offers")
        self.assertEqual(len(rows), 4)
        self.assertEqual(rows[0]["vendor_code"], "B-77")
        self.assertEqual(rows[1]["title"], "Литиевая смазка 400 г")

    def test_pipeline_selects_expected_best_prices_and_writes_xlsx(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            output = Path(directory)
            summary = run(ROOT / "samples", output, ROOT / "config.json")
            self.assertEqual(summary["best_prices"], 3)
            self.assertEqual(summary["mapping_version"], "1.0.0")
            with (output / "best_prices.csv").open(encoding="utf-8-sig", newline="") as handle:
                rows = {row["canonical_id"]: row for row in csv.DictReader(handle)}
            self.assertEqual(rows["bolt-m8x30"]["supplier"], "Supplier B")
            self.assertEqual(rows["lithium-grease"]["supplier"], "Supplier B")
            self.assertEqual(rows["cable-vvg-3x2.5"]["supplier"], "Supplier B")
            xlsx_rows = read_xlsx_rows(output / "price_intelligence.xlsx", "best_prices")
            self.assertEqual(len(xlsx_rows), 3)
            self.assertEqual({row["canonical_id"] for row in xlsx_rows}, set(rows))

    def test_gold_quality_metrics_pass_with_expected_rates(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            output = Path(directory)
            run(ROOT / "samples", output, ROOT / "config.json")
            metrics = json.loads((output / "quality_metrics.json").read_text(encoding="utf-8"))
            self.assertEqual(metrics["precision"], 1.0)
            self.assertEqual(metrics["coverage"], 1.0)
            self.assertEqual(metrics["review_share"], 0.25)
            self.assertTrue(metrics["quality_pass"])

    def test_repeated_import_has_stable_fingerprints_and_no_duplicate_rows(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            output = Path(directory)
            first = run(ROOT / "samples", output, ROOT / "config.json")
            normalized = output / "normalized_offers.csv"
            first_hash = hashlib.sha256(normalized.read_bytes()).hexdigest()
            first_manifest = json.loads((output / "import_manifest.json").read_text(encoding="utf-8"))

            second = run(ROOT / "samples", output, ROOT / "config.json")
            second_hash = hashlib.sha256(normalized.read_bytes()).hexdigest()
            second_manifest = json.loads((output / "import_manifest.json").read_text(encoding="utf-8"))

            self.assertEqual(first["normalized_offers"], 7)
            self.assertEqual(second["normalized_offers"], 7)
            self.assertEqual(second["unchanged_imports"], 2)
            self.assertEqual(first_hash, second_hash)
            self.assertEqual(first_manifest, second_manifest)
            fingerprints = [row["import_fingerprint"] for row in second_manifest["imports"]]
            self.assertEqual(len(fingerprints), len(set(fingerprints)))


if __name__ == "__main__":
    unittest.main()
