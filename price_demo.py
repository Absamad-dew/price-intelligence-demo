"""Normalize supplier price lists and select the best auditable offer."""

from __future__ import annotations

import argparse
import csv
import hashlib
import json
import re
from dataclasses import asdict, dataclass
from difflib import SequenceMatcher
from pathlib import Path
from typing import Dict, Iterable, List, Mapping, Optional, Sequence, Tuple

from xlsx_io import read_xlsx_rows, write_xlsx


UNIT_ALIASES = {
    "шт": ("pc", 1.0),
    "штук": ("pc", 1.0),
    "pc": ("pc", 1.0),
    "pcs": ("pc", 1.0),
    "кг": ("kg", 1.0),
    "kg": ("kg", 1.0),
    "г": ("kg", 0.001),
    "g": ("kg", 0.001),
    "л": ("l", 1.0),
    "l": ("l", 1.0),
    "мл": ("l", 0.001),
    "ml": ("l", 0.001),
    "м": ("m", 1.0),
    "m": ("m", 1.0),
}

STOP_TOKENS = {
    "упаковка",
    "пачка",
    "бухта",
    "метр",
    "метров",
    "шт",
    "штук",
    "кг",
    "г",
}


@dataclass(frozen=True)
class CatalogItem:
    canonical_id: str
    canonical_name: str
    base_unit: str


@dataclass
class NormalizedOffer:
    supplier: str
    supplier_sku: str
    source_name: str
    canonical_id: str
    canonical_name: str
    confidence: float
    quantity_base: float
    base_unit: str
    total_price: float
    price_per_base_unit: float
    currency: str
    source_file: str
    status: str
    mapping_version: str
    import_fingerprint: str


def parse_number(value: str) -> float:
    """Parse a decimal value with comma or dot separators."""
    cleaned = str(value).strip().replace(" ", "").replace(",", ".")
    if not cleaned:
        raise ValueError("empty numeric value")
    return float(cleaned)


def normalize_text(value: str) -> str:
    """Normalize product text for deterministic fuzzy matching."""
    text = value.lower().replace("ё", "е").replace("×", "x").replace("*", "x").replace(",", ".")
    text = re.sub(r"(?<=\d)\s*x\s*(?=\d)", "x", text)
    tokens = re.findall(r"[a-zа-я0-9.]+", text)
    return " ".join(token for token in tokens if token not in STOP_TOKENS)


def match_catalog(name: str, catalog: Iterable[CatalogItem]) -> Tuple[CatalogItem, float]:
    """Return the best catalog match and confidence score."""
    normalized = normalize_text(name)
    best: Optional[CatalogItem] = None
    best_score = -1.0
    for item in catalog:
        candidate = normalize_text(item.canonical_name)
        query_tokens = set(normalized.split())
        candidate_tokens = set(candidate.split())
        overlap = len(query_tokens & candidate_tokens) / max(1, len(candidate_tokens))
        sequence = SequenceMatcher(None, normalized, candidate).ratio()
        score = round(0.65 * overlap + 0.35 * sequence, 4)
        if score > best_score:
            best = item
            best_score = score
    if best is None:
        raise ValueError("catalog is empty")
    return best, best_score


def read_table(path: Path, sheet_name: str | None = None) -> List[Dict[str, str]]:
    """Read a CSV or XLSX flat table."""
    suffix = path.suffix.lower()
    if suffix == ".csv":
        with path.open(encoding="utf-8-sig", newline="") as handle:
            return list(csv.DictReader(handle))
    if suffix == ".xlsx":
        return read_xlsx_rows(path, sheet_name)
    raise ValueError(f"unsupported input format: {path.suffix}")


def load_catalog(path: Path) -> List[CatalogItem]:
    """Load the canonical catalog from CSV or XLSX."""
    return [CatalogItem(**row) for row in read_table(path)]


def import_fingerprint(path: Path, supplier_config: Mapping[str, object], mapping_version: str) -> str:
    """Return a stable content + mapping fingerprint for an import."""
    mapping_payload = json.dumps(
        {"mapping_version": mapping_version, "supplier": supplier_config},
        ensure_ascii=False,
        sort_keys=True,
        separators=(",", ":"),
    ).encode("utf-8")
    digest = hashlib.sha256()
    digest.update(path.read_bytes())
    digest.update(b"\0")
    digest.update(mapping_payload)
    return digest.hexdigest()


def normalize_supplier(
    path: Path,
    supplier_config: Mapping[str, object],
    catalog: List[CatalogItem],
    threshold: float,
    mapping_version: str,
    fingerprint: str,
) -> Tuple[List[NormalizedOffer], List[Dict[str, str]], int]:
    """Normalize one supplier file and return offers, rejected rows, and row count."""
    columns = supplier_config["columns"]
    if not isinstance(columns, dict):
        raise TypeError("supplier columns must be an object")
    supplier = str(supplier_config["supplier"])
    sheet_name = supplier_config.get("sheet")
    rows = read_table(path, None if sheet_name is None else str(sheet_name))
    offers: List[NormalizedOffer] = []
    rejected: List[Dict[str, str]] = []
    for line_number, row in enumerate(rows, start=2):
        sku = row.get(str(columns.get("sku", "")), "").strip()
        source_name = row.get(str(columns.get("name", "")), "").strip()
        try:
            unit_raw = row[str(columns["unit"])].strip().lower()
            if unit_raw not in UNIT_ALIASES:
                raise ValueError(f"unsupported unit: {unit_raw}")
            base_unit, factor = UNIT_ALIASES[unit_raw]
            quantity_base = parse_number(row[str(columns["quantity"])]) * factor
            total_price = parse_number(row[str(columns["price"])])
            if quantity_base <= 0 or total_price < 0:
                raise ValueError("quantity must be positive and price non-negative")
            item, confidence = match_catalog(source_name, catalog)
            status = "matched" if confidence >= threshold and item.base_unit == base_unit else "review"
            offers.append(
                NormalizedOffer(
                    supplier=supplier,
                    supplier_sku=sku,
                    source_name=source_name,
                    canonical_id=item.canonical_id,
                    canonical_name=item.canonical_name,
                    confidence=confidence,
                    quantity_base=round(quantity_base, 6),
                    base_unit=base_unit,
                    total_price=round(total_price, 2),
                    price_per_base_unit=round(total_price / quantity_base, 6),
                    currency=row[str(columns["currency"])].strip(),
                    source_file=path.name,
                    status=status,
                    mapping_version=mapping_version,
                    import_fingerprint=fingerprint,
                )
            )
        except (KeyError, TypeError, ValueError) as error:
            rejected.append(
                {
                    "supplier": supplier,
                    "supplier_sku": sku,
                    "source_file": path.name,
                    "line": str(line_number),
                    "source_name": source_name,
                    "reason": str(error),
                    "mapping_version": mapping_version,
                    "import_fingerprint": fingerprint,
                }
            )
    return offers, rejected, len(rows)


def write_rows(path: Path, rows: Sequence[Mapping[str, object]], fieldnames: Sequence[str]) -> None:
    """Write a stable CSV file."""
    with path.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(rows)


def _round_rate(numerator: int, denominator: int) -> float:
    return round(numerator / denominator, 4) if denominator else 0.0


def calculate_quality_metrics(
    offers: Sequence[NormalizedOffer],
    rejected: Sequence[Mapping[str, str]],
    gold_rows: Sequence[Mapping[str, str]],
    gates: Mapping[str, object],
) -> Dict[str, object]:
    """Evaluate automatic matching against a manually labeled gold sample."""
    predictions: Dict[Tuple[str, str], Tuple[str, str]] = {
        (offer.supplier, offer.supplier_sku): (offer.status, offer.canonical_id) for offer in offers
    }
    predictions.update(
        {(row["supplier"], row["supplier_sku"]): ("review", "") for row in rejected}
    )
    auto_matches = 0
    correct_auto_matches = 0
    matchable_rows = 0
    reviewed_rows = 0
    missing_predictions = 0
    for gold in gold_rows:
        expected = gold.get("expected_canonical_id", "").strip()
        if expected:
            matchable_rows += 1
        prediction = predictions.get((gold["supplier"], gold["supplier_sku"]))
        if prediction is None:
            missing_predictions += 1
            reviewed_rows += 1
            continue
        status, canonical_id = prediction
        if status == "matched":
            auto_matches += 1
            if expected and canonical_id == expected:
                correct_auto_matches += 1
        else:
            reviewed_rows += 1
    precision = _round_rate(correct_auto_matches, auto_matches)
    coverage = _round_rate(correct_auto_matches, matchable_rows)
    review_share = _round_rate(reviewed_rows, len(gold_rows))
    min_precision = float(gates.get("min_precision", 0.0))
    min_coverage = float(gates.get("min_coverage", 0.0))
    max_review_share = float(gates.get("max_review_share", 1.0))
    return {
        "gold_rows": len(gold_rows),
        "matchable_gold_rows": matchable_rows,
        "auto_matched_rows": auto_matches,
        "correct_auto_matches": correct_auto_matches,
        "false_auto_matches": auto_matches - correct_auto_matches,
        "reviewed_or_rejected_rows": reviewed_rows,
        "missing_predictions": missing_predictions,
        "precision": precision,
        "coverage": coverage,
        "review_share": review_share,
        "quality_pass": (
            precision >= min_precision and coverage >= min_coverage and review_share <= max_review_share
        ),
        "gates": {
            "min_precision": min_precision,
            "min_coverage": min_coverage,
            "max_review_share": max_review_share,
        },
    }


def _metric_rows(metrics: Mapping[str, object]) -> List[Dict[str, object]]:
    rows = []
    for key, value in metrics.items():
        if key == "gates":
            for gate, gate_value in dict(value).items():
                rows.append({"metric": f"gate.{gate}", "value": gate_value})
        else:
            rows.append({"metric": key, "value": value})
    return rows


def _load_mapping(config: Mapping[str, object], config_path: Path) -> Tuple[Mapping[str, object], str]:
    mapping_file = config.get("mapping_file")
    if mapping_file:
        mapping = json.loads((config_path.parent / str(mapping_file)).read_text(encoding="utf-8"))
    else:
        mapping = config
    if int(mapping.get("schema_version", 1)) != 1:
        raise ValueError(f"unsupported supplier mapping schema: {mapping.get('schema_version')}")
    mapping_version = str(mapping.get("mapping_version", "legacy-v0"))
    if not isinstance(mapping.get("suppliers"), dict):
        raise TypeError("supplier mapping must contain a suppliers object")
    return mapping, mapping_version


def run(
    input_dir: Path,
    output_dir: Path,
    config_path: Path,
    write_xlsx_output: bool = True,
) -> Dict[str, object]:
    """Run normalization, best-price selection, quality QA, and audited exports."""
    config = json.loads(config_path.read_text(encoding="utf-8"))
    threshold = float(config["confidence_threshold"])
    mapping, mapping_version = _load_mapping(config, config_path)
    catalog_file = str(config.get("catalog_file", "catalog.csv"))
    catalog = load_catalog(input_dir / catalog_file)
    output_dir.mkdir(parents=True, exist_ok=True)
    manifest_path = output_dir / "import_manifest.json"
    previous_records: Dict[str, Mapping[str, object]] = {}
    if manifest_path.exists():
        previous = json.loads(manifest_path.read_text(encoding="utf-8"))
        previous_records = {row["source_file"]: row for row in previous.get("imports", [])}

    offers: List[NormalizedOffer] = []
    rejected: List[Dict[str, str]] = []
    import_records: List[Dict[str, object]] = []
    unchanged_imports = 0
    suppliers = mapping["suppliers"]
    assert isinstance(suppliers, dict)
    for filename, raw_supplier_config in suppliers.items():
        if not isinstance(raw_supplier_config, dict):
            raise TypeError(f"supplier mapping for {filename} must be an object")
        source_path = input_dir / filename
        fingerprint = import_fingerprint(source_path, raw_supplier_config, mapping_version)
        normalized, errors, input_rows = normalize_supplier(
            source_path,
            raw_supplier_config,
            catalog,
            threshold,
            mapping_version,
            fingerprint,
        )
        offers.extend(normalized)
        rejected.extend(errors)
        if previous_records.get(filename, {}).get("import_fingerprint") == fingerprint:
            unchanged_imports += 1
        import_records.append(
            {
                "source_file": filename,
                "supplier": raw_supplier_config["supplier"],
                "mapping_version": mapping_version,
                "import_id": f"sha256:{fingerprint}",
                "import_fingerprint": fingerprint,
                "input_rows": input_rows,
                "normalized_rows": len(normalized),
                "rejected_rows": len(errors),
            }
        )

    offer_rows = [asdict(offer) for offer in offers]
    offer_fields = list(NormalizedOffer.__dataclass_fields__)
    write_rows(output_dir / "normalized_offers.csv", offer_rows, offer_fields)

    matched = [offer for offer in offers if offer.status == "matched"]
    best_by_item: Dict[str, NormalizedOffer] = {}
    for offer in matched:
        current = best_by_item.get(offer.canonical_id)
        if current is None or offer.price_per_base_unit < current.price_per_base_unit:
            best_by_item[offer.canonical_id] = offer
    best_rows = [asdict(best_by_item[key]) for key in sorted(best_by_item)]
    write_rows(output_dir / "best_prices.csv", best_rows, offer_fields)

    review_rows: List[Dict[str, object]] = list(rejected)
    review_rows.extend(
        {
            "supplier": offer.supplier,
            "supplier_sku": offer.supplier_sku,
            "source_file": offer.source_file,
            "line": "",
            "source_name": offer.source_name,
            "reason": f"low confidence or unit mismatch: {offer.confidence}",
            "mapping_version": offer.mapping_version,
            "import_fingerprint": offer.import_fingerprint,
        }
        for offer in offers
        if offer.status == "review"
    )
    review_fields = [
        "supplier",
        "supplier_sku",
        "source_file",
        "line",
        "source_name",
        "reason",
        "mapping_version",
        "import_fingerprint",
    ]
    write_rows(output_dir / "review_queue.csv", review_rows, review_fields)

    gold_file = config.get("gold_sample")
    quality_metrics: Dict[str, object] = {}
    if gold_file:
        gold_rows = read_table(input_dir / str(gold_file))
        quality_metrics = calculate_quality_metrics(
            offers,
            rejected,
            gold_rows,
            config.get("quality_gates", {}),
        )
        (output_dir / "quality_metrics.json").write_text(
            json.dumps(quality_metrics, ensure_ascii=False, indent=2) + "\n",
            encoding="utf-8",
        )

    manifest = {"mapping_version": mapping_version, "imports": import_records}
    manifest_path.write_text(
        json.dumps(manifest, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )
    summary: Dict[str, object] = {
        "input_rows": len(offers) + len(rejected),
        "normalized_offers": len(offers),
        "matched_offers": len(matched),
        "best_prices": len(best_rows),
        "review_items": len(review_rows),
        "imported_files": len(import_records),
        "unchanged_imports": unchanged_imports,
        "mapping_version": mapping_version,
        "quality_pass": quality_metrics.get("quality_pass") if quality_metrics else None,
    }
    (output_dir / "run_summary.json").write_text(
        json.dumps({**summary, "confidence_threshold": threshold}, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )

    if write_xlsx_output:
        workbook_sheets: Dict[str, Sequence[Mapping[str, object]]] = {
            "normalized_offers": offer_rows,
            "best_prices": best_rows,
            "review_queue": review_rows,
            "quality_metrics": _metric_rows(quality_metrics),
            "imports": import_records,
        }
        write_xlsx(output_dir / "price_intelligence.xlsx", workbook_sheets)
    return summary


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--input-dir", type=Path, default=Path("samples"))
    parser.add_argument("--output-dir", type=Path, default=Path("output"))
    parser.add_argument("--config", type=Path, default=Path("config.json"))
    parser.add_argument("--no-xlsx-output", action="store_true")
    args = parser.parse_args()
    print(
        json.dumps(
            run(args.input_dir, args.output_dir, args.config, not args.no_xlsx_output),
            ensure_ascii=False,
        )
    )


if __name__ == "__main__":
    main()
