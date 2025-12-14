"""
fleet_reconcile.py

Fleet master-data reconciliation:
- Compares HR cost center list vs Fleet export (drivers/cars/cost centers)
- Outputs 3 Excel files:
  1) missing_in_hr.xlsx          (drivers with cars in fleet but no match in HR list)
  2) costcenter_mismatch.xlsx    (matched drivers with different cost center HR vs Fleet)
  3) fleet_mapping_refreshed.xlsx (license plate -> driver -> cost center mapping for invoice export)

USAGE:
  python src/fleet_reconcile.py \
    --fleet "FleetManagerExport.xlsx" \
    --hr "Aktive_Mitarbeiter_mit_Stammkostenstelle.xlsx" \
    --out-dir "outputs" \
    --threshold 95 \
    --exclude-pool

Notes:
- This script expects *synthetic* data in public repos. Do NOT upload employer data.
"""

from __future__ import annotations

import argparse
from pathlib import Path
from typing import Optional, Tuple

import pandas as pd
from fuzzywuzzy import fuzz, process


# -----------------------------
# Column configuration (expected)
# -----------------------------
FLEET_COL_FIRST = "First name"
FLEET_COL_LAST = "Name"
FLEET_COL_LICENSES = "License Numbers"
FLEET_COL_COSTCENTER = "Cost center"

HR_COL_FIRST = "Vorname"
HR_COL_LAST = "Nachname"
HR_COL_COSTCENTER = "Kostenstelle"
HR_COL_CC_DESC = "Bezeichnung d.KST"


# -----------------------------
# Helpers
# -----------------------------
def _require_columns(df: pd.DataFrame, required: list[str], label: str) -> None:
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError(
            f"[{label}] Missing required columns: {missing}\n"
            f"Available columns: {list(df.columns)}"
        )


def _clean_str_series(s: pd.Series) -> pd.Series:
    return s.astype(str).str.strip().replace({"nan": pd.NA, "None": pd.NA, "": pd.NA})


def _build_fullname(df: pd.DataFrame, first_col: str, last_col: str, out_col: str = "FullName") -> pd.DataFrame:
    df = df.copy()
    df[first_col] = _clean_str_series(df[first_col])
    df[last_col] = _clean_str_series(df[last_col])
    df[out_col] = (df[first_col].fillna("") + " " + df[last_col].fillna("")).str.strip()
    df.loc[df[out_col].eq(""), out_col] = pd.NA
    return df


def _best_fuzzy_match(name: Optional[str], choices: list[str], threshold: int) -> Tuple[Optional[str], int]:
    if name is None:
        return None, 0
    name_str = str(name).strip()
    if not name_str or name_str.lower() == "nan":
        return None, 0

    best = process.extractOne(name_str, choices, scorer=fuzz.token_sort_ratio)
    if not best:
        return None, 0

    best_name, score = best[0], int(best[1])
    if score >= threshold:
        return best_name, score
    return None, score


def _has_any_car(licenses_value: object) -> bool:
    if pd.isna(licenses_value):
        return False
    s = str(licenses_value).strip()
    return s != "" and s.lower() != "nan"


def _split_and_clean_plates(licenses_value: object) -> list[str]:
    """
    Fleet exports sometimes contain multiple plates separated by commas.
    We clean spaces and hyphens, and uppercase.
    """
    if pd.isna(licenses_value):
        return []
    raw = str(licenses_value)
    plates = []
    for p in raw.split(","):
        p = p.strip()
        if not p:
            continue
        cleaned = p.replace(" ", "").replace("-", "").upper()
        if cleaned:
            plates.append(cleaned)
    return plates


# -----------------------------
# Core logic
# -----------------------------
def run_reconciliation(
    fleet_path: Path,
    hr_path: Path,
    out_dir: Path,
    threshold: int = 95,
    exclude_pool: bool = True,
) -> None:
    out_dir.mkdir(parents=True, exist_ok=True)

    # Load
    fleet_df = pd.read_excel(fleet_path)
    hr_df = pd.read_excel(hr_path)

    # Validate columns
    _require_columns(
        fleet_df,
        [FLEET_COL_FIRST, FLEET_COL_LAST, FLEET_COL_LICENSES, FLEET_COL_COSTCENTER],
        label="FLEET",
    )
    _require_columns(
        hr_df,
        [HR_COL_FIRST, HR_COL_LAST, HR_COL_COSTCENTER],
        label="HR",
    )
    # HR description is optional (nice to have)
    has_hr_desc = HR_COL_CC_DESC in hr_df.columns

    # Build FullName
    fleet_df = _build_fullname(fleet_df, FLEET_COL_FIRST, FLEET_COL_LAST, out_col="FullName")
    hr_df = _build_fullname(hr_df, HR_COL_FIRST, HR_COL_LAST, out_col="FullName")

    # Prepare choices (HR full names)
    hr_names = hr_df["FullName"].dropna().astype(str).unique().tolist()

    # Fuzzy match each fleet name to HR
    matched_hr_name = []
    match_score = []
    for nm in fleet_df["FullName"].tolist():
        best_name, score = _best_fuzzy_match(nm, hr_names, threshold=threshold)
        matched_hr_name.append(best_name)
        match_score.append(score)

    fleet_df["HR_MatchName"] = matched_hr_name
    fleet_df["HR_MatchScore"] = match_score
    fleet_df["Matched"] = fleet_df["HR_MatchName"].notna()

    # -----------------------------
    # Output 1: missing_in_hr.xlsx
    # -----------------------------
    # Only unmatched + has car + (optional) exclude pool/poolwagen
    mask_unmatched = ~fleet_df["Matched"]
    mask_has_car = fleet_df[FLEET_COL_LICENSES].apply(_has_any_car)

    if exclude_pool:
        mask_not_pool = ~fleet_df["FullName"].astype(str).str.contains("pool", case=False, na=False)
        missing_in_hr = fleet_df[mask_unmatched & mask_has_car & mask_not_pool].copy()
    else:
        missing_in_hr = fleet_df[mask_unmatched & mask_has_car].copy()

    missing_in_hr_out = missing_in_hr[[FLEET_COL_FIRST, FLEET_COL_LAST, "FullName", FLEET_COL_LICENSES]].copy()
    missing_file = out_dir / "missing_in_hr.xlsx"
    missing_in_hr_out.to_excel(missing_file, index=False)

    # -----------------------------
    # Output 2: costcenter_mismatch.xlsx
    # -----------------------------
    # Build HR lookup based on HR FullName
    hr_kst_lookup = hr_df.set_index("FullName")[HR_COL_COSTCENTER].to_dict()
    hr_desc_lookup = hr_df.set_index("FullName")[HR_COL_CC_DESC].to_dict() if has_hr_desc else {}

    matched_df = fleet_df[fleet_df["Matched"]].copy()

    # Attach HR cost center/desc using the matched HR name (fuzzy-approved)
    matched_df["HR_Kostenstelle"] = matched_df["HR_MatchName"].map(hr_kst_lookup)
    if has_hr_desc:
        matched_df["HR_Bezeichnung_d_KST"] = matched_df["HR_MatchName"].map(hr_desc_lookup)
    else:
        matched_df["HR_Bezeichnung_d_KST"] = pd.NA

    # Compare numerically to avoid '27100.0' vs '27100'
    matched_df["Fleet_CC_num"] = pd.to_numeric(matched_df[FLEET_COL_COSTCENTER], errors="coerce")
    matched_df["HR_CC_num"] = pd.to_numeric(matched_df["HR_Kostenstelle"], errors="coerce")

    cond_both_diff = matched_df["Fleet_CC_num"].notna() & matched_df["HR_CC_num"].notna() & (matched_df["Fleet_CC_num"] != matched_df["HR_CC_num"])
    cond_fleet_missing = matched_df["Fleet_CC_num"].isna() & matched_df["HR_CC_num"].notna()
    cond_hr_missing = matched_df["Fleet_CC_num"].notna() & matched_df["HR_CC_num"].isna()

    mismatch_df = matched_df[cond_both_diff | cond_fleet_missing | cond_hr_missing].copy()

    mismatch_out = mismatch_df[
        [
            "FullName",
            "HR_MatchName",
            "HR_MatchScore",
            FLEET_COL_COSTCENTER,
            "HR_Kostenstelle",
            "HR_Bezeichnung_d_KST",
            FLEET_COL_LICENSES,
        ]
    ].copy()

    mismatch_file = out_dir / "costcenter_mismatch.xlsx"
    mismatch_out.to_excel(mismatch_file, index=False)

    # -----------------------------
    # Output 3: fleet_mapping_refreshed.xlsx
    # -----------------------------
    # One row per plate (cleaned), with Fleet CC and driver name
    rows = []
    for _, row in fleet_df.iterrows():
        plates = _split_and_clean_plates(row.get(FLEET_COL_LICENSES))
        if not plates:
            continue
        for plate in plates:
            rows.append(
                {
                    "License Number": plate,
                    "Cost center": row.get(FLEET_COL_COSTCENTER),
                    "FullName": row.get("FullName"),
                }
            )

    mapping_df = pd.DataFrame(rows, columns=["License Number", "Cost center", "FullName"])

    # Optional cleanup: normalize cost center to numeric if possible
    mapping_df["Cost center"] = pd.to_numeric(mapping_df["Cost center"], errors="ignore")

    mapping_file = out_dir / "fleet_mapping_refreshed.xlsx"
    mapping_df.to_excel(mapping_file, index=False)

    # Print a short summary
    print("Done.")
    print(f"- missing_in_hr:       {len(missing_in_hr_out)} rows -> {missing_file}")
    print(f"- costcenter_mismatch: {len(mismatch_out)} rows -> {mismatch_file}")
    print(f"- refreshed_mapping:   {len(mapping_df)} rows -> {mapping_file}")


# -----------------------------
# CLI
# -----------------------------
def main() -> None:
    parser = argparse.ArgumentParser(description="Fleet master-data reconciliation (HR vs Fleet) -> 3 Excel outputs")
    parser.add_argument("--fleet", required=True, type=str, help="Path to fleet export Excel file")
    parser.add_argument("--hr", required=True, type=str, help="Path to HR cost center Excel file")
    parser.add_argument("--out-dir", default="outputs", type=str, help="Output directory (default: outputs)")
    parser.add_argument("--threshold", default=95, type=int, help="Fuzzy match threshold (default: 95)")
    parser.add_argument(
        "--exclude-pool",
        action="store_true",
        help="Exclude names containing 'pool' from missing_in_hr output",
    )

    args = parser.parse_args()

    run_reconciliation(
        fleet_path=Path(args.fleet),
        hr_path=Path(args.hr),
        out_dir=Path(args.out_dir),
        threshold=int(args.threshold),
        exclude_pool=bool(args.exclude_pool),
    )


if __name__ == "__main__":
    main()

