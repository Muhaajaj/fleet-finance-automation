# Fleet Finance Automation

Python tools to keep fleet cost-center data consistent and to generate a booking-ready invoice export.

## What’s inside

### 1) Fleet master data reconciliation
Compares:
- HR cost center list (names + cost centers)
- Fleet system list (drivers + cars + cost centers)

Outputs 3 Excel files:
- **missing_in_hr.xlsx**: drivers with allocated cars but no match in HR list (new driver or name mismatch)
- **costcenter_mismatch.xlsx**: drivers found in both sources but with different cost centers (department change)
- **fleet_mapping_refreshed.xlsx**: license plate → driver → cost center (input for step 2)

Ideal result: the first two files are empty, and the refreshed mapping is complete.

### 2) Invoice export (booking-ready CSV)
Takes:
- mapping file from step 1
- invoice detail CSV (vendor export)

Outputs:
- booking-ready CSV for import into an accounting system

Validation:
- If any invoice line has no cost center match, the script exports **missing_costcenters.xlsx** and stops.

## Data privacy
All sample files in this repo are **synthetic**. Do NOT upload real employer data, invoices, license plates, names, cost centers, or account numbers.

## Tools
Python, pandas, openpyxl, NumPy

## Quick demo (synthetic)
- Sample inputs: `data/sample/`
- Example validation output (missing mapping): `docs/examples/missing_costcenters_example.csv`
