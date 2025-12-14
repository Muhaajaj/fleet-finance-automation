"""
invoice_export.py

Invoice export:
- Reads an invoice detail CSV (vendor export)
- Reads the refreshed fleet mapping file (from fleet_reconcile.py):
    fleet_mapping_refreshed.xlsx with columns:
      - License Number
      - Cost center
      - FullName
- Produces a booking-ready CSV for import.
- Validation gate:
    If any invoice line has no matching cost center (by license plate),
    exports missing_costcenters.xlsx and stops.

USAGE:
  python src/invoice_export.py \
    --invoice "DKV_detail.csv" \
    --mapping "outputs/fleet_mapping_refreshed.xlsx" \
    --out "outputs/invoice_booking_export.csv" \
    --missing-out "outputs/missing_costcenters.xlsx"

Notes:
- Public repos must use synthetic sample data only.
"""

from __future__ import annotations

import argparse
from pathlib import Path
from typing import Optional

import pandas as pd


# -----------------------------
# Helpers
# -----------------------------
def _clean_currency_de(value: object) -> float:
    """
    Converts German-formatted numbers like '1.234,56' to float 1234.56
    """
    if pd.isna(value):
        return float("nan")
    s = str(value).strip()
    if not s or s.lower() == "nan":
        return float("nan")
    s = s.replace(".", "").replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return float("nan")


def _standardize_plate(value: object) -> Optional[str]:
    """
    Removes spaces and hyphens, uppercases. Returns None if empty.
    """
    if pd.isna(value):
        return None
    s = str(value).strip()
    if not s or s.lower() == "nan":
        return None
    s = s.replace(" ", "").replace("-", "").upper()
    return s or None


def _infer_first_matching_col(columns: list[str], contains_text: str) -> str:
    for c in columns:
        if contains_text.lower() in str(c).lower():
            return c
    raise ValueError(f"Could not find a column containing '{contains_text}'. Columns: {columns}")


# -----------------------------
# Core
# -----------------------------
def run_invoice_export(
    invoice_csv: Path,
    mapping_xlsx: Path,
    out_csv: Path,
    missing_out_xlsx: Path,
    encoding: str = "latin1",
    sep: str = ";",
) -> None:
    # Load invoice - read full file (for header fields) and detail area
    full = pd.read_csv(invoice_csv, encoding=encoding, sep=sep)
    detail = pd.read_csv(invoice_csv, encoding=encoding, sep=sep, skiprows=range(1, 5))

    # Mapping file from fleet_reconcile.py
    mapping = pd.read_excel(mapping_xlsx)

    # Validate mapping columns
    required_map_cols = {"License Number", "Cost center"}
    if not required_map_cols.issubset(set(mapping.columns)):
        raise ValueError(f"Mapping file must contain {required_map_cols}. Found: {list(mapping.columns)}")

    # Identify relevant columns in invoice
    # Your original invoice column names included:
    # - "Kennzeichen"
    # - "Wert incl. USt.1"
    # - "USt"
    plate_col = "Kennzeichen" if "Kennzeichen" in detail.columns else _infer_first_matching_col(list(detail.columns), "kennzeichen")
    gross_col = "Wert incl. USt.1" if "Wert incl. USt.1" in detail.columns else _infer_first_matching_col(list(detail.columns), "wert")
    vat_col = "USt" if "USt" in detail.columns else _infer_first_matching_col(list(detail.columns), "ust")

    # Clean invoice values
    detail = detail.copy()
    detail[gross_col] = detail[gross_col].apply(_clean_currency_de)
    detail["Plate_std"] = detail[plate_col].apply(_standardize_plate)

    # Drop rows with no plate (cannot allocate)
    detail = detail[detail["Plate_std"].notna()].copy()

    # Build VAT code: 19% -> 9 else 50 (same as your logic)
    detail["Steuerschluessel"] = detail[vat_col].astype(str).str.strip().eq("19%").map({True: 9, False: 50})

    # Prepare mapping lookup
    mapping = mapping.copy()
    mapping["Plate_std"] = mapping["License Number"].apply(_standardize_plate)
    mapping["Cost center"] = pd.to_numeric(mapping["Cost center"], errors="coerce")

    # Deduplicate mapping by plate (first occurrence)
    mapping = mapping.dropna(subset=["Plate_std"]).drop_duplicates(subset=["Plate_std"], keep="first")

    # Join invoice lines to mapping
    merged = detail.merge(mapping[["Plate_std", "Cost center"]], on="Plate_std", how="left")

    # Missing cost centers -> export and stop
    missing = merged[merged["Cost center"].isna()][["Plate_std"]].drop_duplicates().sort_values("Plate_std")
    if len(missing) > 0:
        missing.rename(columns={"Plate_std": "License Number (standardized)"}, inplace=True)
        missing_out_xlsx.parent.mkdir(parents=True, exist_ok=True)
        missing.to_excel(missing_out_xlsx, index=False)
        raise SystemExit(
            f"Missing cost center mapping for {len(missing)} license plates. "
            f"Exported: {missing_out_xlsx}. Please update mapping and rerun."
        )

    # Try to infer invoice date & invoice number from the full file
    # (kept intentionally robust; if missing, user can still use the export)
    invoice_date = None
    try:
        invoice_date = full.iloc[0, 1]
    except Exception:
        invoice_date = None

    invoice_date_dt = pd.to_datetime(invoice_date, dayfirst=True, errors="coerce")
    date_str = invoice_date_dt.strftime("%d.%m.%Y") if pd.notna(invoice_date_dt) else ""

    # Create a simple "Belegnr" if a "Rechnung" column exists
    belegnr = ""
    try:
        rechnung_col = _infer_first_matching_col(list(full.columns), "rechnung")
        base_rechnung = "".join([ch for ch in str(full[rechnung_col].dropna().iloc[0]) if ch.isdigit()])[:9]
        year_suffix = f"{(invoice_date_dt.year % 100):02d}" if pd.notna(invoice_date_dt) else ""
        belegnr = f"{year_suffix}/{base_rechnung}" if year_suffix and base_rechnung else base_rechnung
    except Exception:
        belegnr = ""

    # Build booking export table (structure similar to your output)
    merged["Betrag"] = pd.to_numeric(merged[gross_col], errors="coerce").round(2)
    merged["Cost center"] = merged["Cost center"].astype("Int64")

    summary_row = {
        "Buchungsdatum": date_str,
        "Belegdatum": date_str,
        "Belegnr.": belegnr,
        "Gegenkonto": "80071244",
        "Steuerschlüssel": "",
        "Kontonr.": "",
        "Beschreibung": f"Fleet invoice {date_str}".strip(),
        "": "",
        "Betrag": -merged["Betrag"].sum(),
        "KostenstelleCode": "",
    }

    details = pd.DataFrame(
        {
            "Buchungsdatum": date_str,
            "Belegdatum": date_str,
            "Belegnr.": belegnr,
            "Gegenkonto": "80071244",
            "Steuerschlüssel": merged["Steuerschluessel"],
            "Kontonr.": "4530",
            "Beschreibung": merged["Plate_std"],
            "": "",
            "Betrag": merged["Betrag"],
            "KostenstelleCode": merged["Cost center"],
        }
    )

    export_df = pd.concat([pd.DataFrame([summary_row]), details], ignore_index=True)

    out_csv.parent.mkdir(parents=True, exist_ok=True)
    with open(out_csv, "w", encoding=encoding, newline="") as f:
        f.write("BearbeitenFibuBuch.BlattDKV\n")
        export_df.to_csv(f, index=False, sep=";", decimal=",", na_rep="", float_format="%.2f")

    print(f"Done. Exported booking CSV: {out_csv}")


def main() -> None:
    parser = argparse.ArgumentParser(description="Generate booking-ready invoice export using fleet mapping")
    parser.add_argument("--invoice", required=True, type=str, help="Path to invoice detail CSV")
    parser.add_argument("--mapping", required=True, type=str, help="Path to fleet_mapping_refreshed.xlsx")
    parser.add_argument("--out", default="outputs/invoice_booking_export.csv", type=str, help="Output CSV path")
    parser.add_argument("--missing-out", default="outputs/missing_costcenters.xlsx", type=str, help="Missing mapping output XLSX path")
    parser.add_argument("--encoding", default="latin1", type=str, help="CSV encoding (default: latin1)")
    parser.add_argument("--sep", default=";", type=str, help="CSV separator (default: ;)")

    args = parser.parse_args()

    run_invoice_export(
        invoice_csv=Path(args.invoice),
        mapping_xlsx=Path(args.mapping),
        out_csv=Path(args.out),
        missing_out_xlsx=Path(args.missing_out),
        encoding=args.encoding,
        sep=args.sep,
    )


if __name__ == "__main__":
    main()

