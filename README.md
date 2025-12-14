# Fleet Finance Automation

Python tools to keep fleet cost-center data consistent and to generate a booking-ready invoice export.

## What’s inside

### 1) Fleet master data reconciliation (`src/fleet_reconcile.py`)
Compares:
- HR cost center list (names + cost centers)
- Fleet export (drivers + license plates + cost centers)

Outputs:
- `missing_in_hr.xlsx` — drivers with assigned cars in fleet but no match in HR list (new driver / name mismatch)
- `costcenter_mismatch.xlsx` — drivers found in both sources but with different cost centers (department change)
- `fleet_mapping_refreshed.xlsx` — license plate → driver → cost center (input for step 2)

Ideal result: the first two outputs are empty, and the refreshed mapping is complete.

### 2) Invoice export (`src/invoice_export.py`)
Takes:
- mapping file from step 1 (`fleet_mapping_refreshed.xlsx`)
- invoice detail CSV export

Outputs:
- booking-ready CSV for import
- validation output: `missing_costcenters.xlsx` (only if any invoice line has no cost center match)

Validation rule:
- If any invoice line has no cost center match, the script exports `missing_costcenters.xlsx` and stops.

## Quick demo (synthetic)
- Sample inputs are in: `data/sample/`
- Colab demo notebook: `fleet_finance_automation_demo.ipynb`

## Data privacy
All sample files in this repo are synthetic.
Do NOT upload real employer data, invoices, license plates, names, cost centers, account numbers, or internal identifiers.

## Tools
Python, pandas, openpyxl, NumPy
