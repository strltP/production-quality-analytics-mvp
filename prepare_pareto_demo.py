"""Convert the bundled public synthetic pass/fail CSV to the input schema."""

from pathlib import Path

import pandas as pd


def prepare(source: Path, destination: Path):
    """One upstream inspection ID represents one simulated inspected unit."""
    data = pd.read_csv(source, dtype=str, keep_default_na=False)
    required = {"inspection_id", "date", "production_line", "product_model", "status", "defect_type"}
    if required - set(data.columns):
        raise ValueError("Missing fields in the upstream Pareto demo CSV.")
    if data["inspection_id"].eq("").any() or data["inspection_id"].duplicated().any():
        raise ValueError("The upstream inspection IDs must be present and unique.")
    if not data["status"].isin(["Pass", "Fail"]).all():
        raise ValueError("Expected Pass or Fail in the upstream status field.")
    failed = data["status"].eq("Fail")
    if not failed.equals(data["defect_type"].ne("")):
        raise ValueError("Expected exactly the failed inspections to have a defect type.")
    prepared = data[["inspection_id", "date", "production_line", "product_model", "defect_type"]].copy()
    prepared["inspected_units"] = 1
    prepared["defective_units"] = failed.astype(int)
    destination.parent.mkdir(parents=True, exist_ok=True)
    prepared.to_csv(destination, index=False, encoding="utf-8-sig")
    return len(prepared)


if __name__ == "__main__":
    root = Path(__file__).resolve().parent
    output = root / "data/pareto_input/public_synthetic_inspections.csv"
    count = prepare(root / "data/raw/quality_inspections.csv", output)
    print(f"Prepared {count:,} records in {output}")
