# MUKERC — Multi-Use Kilowatt-hour Electricity Rate Calculator

A Python pipeline for decomposing UK non-domestic electricity costs across all network tariff components. Covers all 14 GB DNO regions and financial years 2017–2027, with support for Generic benchmark analysis and Specific site-level case studies.

**Cost components modelled:**
- Distribution Use of System (DUoS) — volumetric, fixed, capacity, reactive, and exceeded capacity charges
- Transmission Network Use of System (TNUoS) — locational and residual components, including the 2023 Targeted Charging Review (TCR) reform
- Balancing Services Use of System (BSUoS)
- Policy levies — RO, CFD, CM, WHD, and other applicable charges
- Wholesale commodity — half-hourly imbalance price data from ELEXON MID
- Miscellaneous charges

## Quick Start

```bash
pip install pandas numpy matplotlib openpyxl
```

1. Open `Code Run Folder/Manual_Input.py` and set your run parameters.
2. **Generic mode only:** run `Build Price Profiles.py` first to generate the flat-load price profiles that Post Analysis reads.
3. Run `Post Analysis.py`.

Outputs are written to `Output Data Repository/` (created automatically on first run).

For full configuration and input data reference, see [USER_GUIDE.md](USER_GUIDE.md).

## Modes

| Mode | Use case | Scripts to run |
|------|----------|----------------|
| **Generic** | Benchmark costs across all DNOs/voltages/bands using representative flat loads | `Build Price Profiles.py` → `Post Analysis.py` |
| **Specific** | Case study for a real site with a measured demand profile | `Post Analysis.py` only |

## Repository Structure

```
Code Run Folder/
├── Manual_Input.py                  # Script 1 — All user settings (edit before each run)
├── Helper_Functions.py              # Script 2 — Shared utilities
├── Build Price Profiles.py          # Script 3 — Builds Generic flat-load price profiles
├── Post Analysis.py                 # Script 4 — Financial analysis and output generation
├── Modules/
│   ├── Helper_Functions.py          # Module copy (imported by scripts at runtime)
│   └── Case_Study_Analytics.py      # Case study helper functions
├── Input/
│   ├── Network Prices/
│   │   ├── Distribution/{year}/     # DNO tariff workbooks (14 regions × 11 years)
│   │   ├── Transmission/            # TNUoS Rates.xlsx
│   │   └── Balancing/               # BSUoS Rates.xlsx
│   ├── Policy Levy Prices/          # Policy Levy Rates.xlsx
│   ├── Commodity Prices/            # ELEXON MID half-hourly price CSVs
│   ├── Miscellaneous Prices/        # Miscellaneous Rates.xlsx
│   └── Case Study Demand Profiles/  # Half-hourly demand profile workbooks (Specific mode)
└── Case Study Supporting Files/
    ├── Case_Study_Parameters.py     # Reference document — case study run parameters
    └── Generate_TOU_Shift_Profiles.py  # Generates TOU-shifted demand profiles

Output Data Repository/
├── Generic Results/                 # Generic mode outputs by voltage type and band
└── Case Studies/                    # Specific mode outputs by case study name
```

## Prerequisites

- Python 3.9 or later
- pandas, numpy, matplotlib, openpyxl

```bash
pip install pandas numpy matplotlib openpyxl
```

## Authors

Luke Bennett — Imperial College London
