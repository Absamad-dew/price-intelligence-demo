# System analysis case: supplier price reconciliation

This case describes the behavior that is executable in this repository and a
minimal boundary for turning it into a service. The API section is a proposed
contract, not an implemented or production-tested endpoint.

## Actors

- **Procurement specialist** — uploads price lists, reviews uncertain rows and
  consumes the best-price result.
- **Data/operations administrator** — maintains the canonical catalog,
  supplier column mappings, matching threshold and quality gates.
- **Supplier** — external source of CSV/XLSX offers with its own schema, units
  and naming conventions.
- **Downstream purchasing system** — future consumer of approved results;
  Google Sheets, ERP and database integrations are outside the current demo.

## AS-IS pain

Supplier files use different column names, product wording, pack sizes and
units. Manual consolidation makes it difficult to compare unit prices, explain
why two rows were matched, separate uncertain rows from safe automatic matches,
and prove which source file and mapping produced a decision.

## TO-BE flow demonstrated here

1. Read a canonical catalog, supplier CSV/XLSX files, versioned mappings and
   configuration.
2. Validate and normalize supplier rows to one schema and base units.
3. Fingerprint each import from source bytes, canonical mapping and mapping
   version.
4. Match offers to catalog products with deterministic rules and a configured
   confidence threshold.
5. Route malformed, low-confidence or unit-incompatible rows to a review queue.
6. Calculate price per base unit and select the lowest comparable offer per
   canonical product.
7. Evaluate the run against a labeled acceptance sample and write CSV, JSON and
   XLSX artifacts with provenance.

## Functional requirements

- **FR-01:** accept ordinary flat CSV and XLSX supplier tables in one run.
- **FR-02:** map supplier-specific columns through a versioned configuration.
- **FR-03:** normalize supported quantities and units before price comparison.
- **FR-04:** preserve source file, supplier SKU, mapping version and import
  fingerprint on normalized decisions.
- **FR-05:** never silently discard a malformed or uncertain row; record it in
  `review_queue.csv` with a reason.
- **FR-06:** select best prices only from rows classified as matched.
- **FR-07:** publish run counts, quality metrics and the quality-gate result.
- **FR-08:** produce machine-readable CSV/JSON outputs and a five-sheet XLSX
  workbook.

## Non-functional requirements

- **Reproducibility:** identical source bytes and mapping version produce stable
  fingerprints and deterministic outputs.
- **Auditability:** every decision retains source and mapping provenance.
- **Safety:** low-confidence decisions require review instead of forced matching.
- **Portability:** the core demo runs offline on Python 3.10+ without an external
  model or API; XLSX interoperability is checked independently with `openpyxl`.
- **Testability:** quality gates are explicit configuration, not hidden business
  logic.

## Minimal proposed API/JSON contract

The current interface is CLI/filesystem-based. A service wrapper could preserve
the same semantics without changing the decision engine:

```http
POST /v1/import-runs
Content-Type: multipart/form-data

files=<one or more CSV/XLSX files>
options={"mapping_version":"1.0.0","confidence_threshold":0.58}
```

```json
{
  "run_id": "run_01",
  "status": "completed",
  "mapping_version": "1.0.0",
  "counts": {
    "input_rows": 8,
    "normalized_offers": 7,
    "matched_offers": 6,
    "best_prices": 3,
    "review_items": 2
  },
  "quality": {
    "precision": 1.0,
    "coverage": 1.0,
    "review_share": 0.25,
    "quality_pass": true
  },
  "artifacts": {
    "best_prices": "/v1/import-runs/run_01/best-prices",
    "review_queue": "/v1/import-runs/run_01/review-queue"
  }
}
```

For production, the contract would also require authentication, upload limits,
idempotency keys, durable run states, error schemas and artifact retention.

## Acceptance criteria on the versioned demo sample

- One command processes both `supplier_a.csv` and a real `supplier_b.xlsx`.
- Eight input rows produce seven normalized offers, six automatic matches,
  three best-price rows and two review items.
- The missing-price row and the low-confidence row appear in the review queue
  with explicit reasons.
- The labeled sample passes gates: precision `1.0`, coverage `1.0`, review share
  `0.25`.
- A repeated run keeps import fingerprints stable and does not duplicate offers.
- The generated workbook can be opened by `openpyxl` and contains the expected
  sheets, structure and formulas.

Run the evidence:

```powershell
python price_demo.py
python -m unittest discover -s tests -v
```

## Risks and non-claims

- The eight-row labeled sample proves the pipeline and scorer, not production
  matching quality; a client acceptance set is still required.
- The XLSX adapter supports flat tables, not arbitrary formulas, merged cells,
  PDFs or multi-section documents.
- The matching vocabulary, unit rules, thresholds and quality gates require
  domain validation before use on purchasing decisions.
- Currency conversion, Google Sheets/ERP integration, PostgreSQL, authentication,
  concurrent uploads and a human-review UI are not implemented.
- No LLM is used, and this repository does not claim commercial deployment,
  production scale, or previous operation on client data.
