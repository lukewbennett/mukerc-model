# MUKERC User Guide

## Contents

1. [Overview](#1-overview)
2. [Setup](#2-setup)
3. [Running Generic Mode](#3-running-generic-mode)
4. [Running Specific Mode](#4-running-specific-mode)
5. [Manual_Input.py Reference](#5-manual_inputpy-reference)
6. [Input Data Reference](#6-input-data-reference)
7. [Output Structure](#7-output-structure)
8. [Extending to New Financial Years](#8-extending-to-new-financial-years)
9. [Case Study Supporting Files](#9-case-study-supporting-files)

---

## 1. Overview

MUKERC is a four-script Python pipeline that models UK non-domestic electricity costs at half-hourly resolution, then aggregates them into annual and quarterly financial breakdowns.

### Pipeline scripts

| # | Script | Role |
|---|--------|------|
| 1 | `Manual_Input.py` | All user-configurable settings — edit before each run |
| 2 | `Helper_Functions.py` / `Modules/Helper_Functions.py` | Shared utility functions (imported, never run directly) |
| 3 | `Build Price Profiles.py` | Builds flat-load price profiles for Generic mode |
| 4 | `Post Analysis.py` | Financial analysis and output workbook generation |

### Two operating modes

**Generic mode** sweeps all 14 GB DNO regions, 4 voltage connection types, and 4 residual charging bands using representative flat demand levels per band. It produces a benchmark cost matrix that can be compared across any combination of region, voltage, and band.

**Specific mode** analyses a named case study for one real or representative site. It reads a measured half-hourly demand profile (an `.xlsx` workbook) and one or more DNO regions, applying the full cost model to that specific load shape across each financial year in the configured range.

---

## 2. Setup

### Python version

Python 3.9 or later.

### Required packages

```bash
pip install pandas numpy matplotlib openpyxl
```

All four pipeline scripts use only these packages plus the Python standard library.

### Directory layout

Run all scripts from the `Code Run Folder/` directory, or open them in an IDE with `Code Run Folder/` as the working directory. The scripts resolve all input and output paths relative to their own location, so the working directory at launch does not matter as long as the scripts are run from within `Code Run Folder/`.

---

## 3. Running Generic Mode

Generic mode produces a cost breakdown for a representative flat-load consumer at each combination of DNO region, voltage type, and residual charging band.

### Step 1 — Configure Manual_Input.py

Open `Manual_Input.py` and set:

```python
run_mode = 'Generic'
start_year     = 2017   # first financial year to include
end_year_limit = 2026   # last financial year to include
```

Optionally adjust the representative flat loads in **Section 4** (`generic_flat_loads`) if you want to use different demand levels per band.

To run a quick test on a reduced scope before the full run, set:

```python
build_scope_mode  = 'Test'
build_test_vol_types   = ['HV']
build_test_dno_regions = ['12 - London.xlsx']
build_test_bands       = ['Band 4']
```

### Step 2 — Run Build Price Profiles.py

```bash
python "Build Price Profiles.py"
```

This generates half-hourly price profiles for each DNO region, voltage type, and band combination, writing intermediate files consumed by Post Analysis in Generic mode. Runtime is approximately 30–60 minutes for the full scope.

### Step 3 — Run Post Analysis.py

```bash
python "Post Analysis.py"
```

Post Analysis reads the price profiles produced in Step 2, applies the full cost model, and writes the output workbooks to `Output Data Repository/Generic Results/`.

---

## 4. Running Specific Mode

Specific mode requires only Post Analysis — Build Price Profiles is not needed. It reads Generic mode's price profiles internally and applies them to the supplied demand profile.

### Step 1 — Prepare the demand profile

Place a half-hourly demand profile workbook in:

```
Code Run Folder/Input/Case Study Demand Profiles/
```

The workbook must follow the proforma format (`Demand Profile Proforma.xlsx` in the same folder): column headers at row 5 (0-indexed row 4), a `Date` column containing datetime values, and an `Overall` column containing half-hourly demand in kW.

If you want to analyse a flat demand level rather than a measured profile, set `demand_profile_file = 'Flat'` and configure `basis` with the desired kW level.

### Step 2 — Configure Manual_Input.py

Open `Manual_Input.py` and set:

```python
run_mode = 'Specific'
```

Configure **Section 3**:

```python
case_study_name     = 'My Site - Scenario A'   # creates subfolder under Case Studies/
specific_regions    = ['12 - London.xlsx']       # one or more DNO regions
demand_profile_file = 'My Site Demand Profile.xlsx'  # or 'Flat'
vol_type            = 'HV'                       # connection voltage type
has_mic             = True                       # True if site has a Metering Import Capacity
residual_band       = 'Band 4'                   # TCR residual charging band
basis               = 16500                      # kW — flat demand level or TRIAD demand basis
power_factor        = 0.98
```

Set `start_year` and `end_year_limit` as required.

### Step 3 — Run Post Analysis.py

```bash
python "Post Analysis.py"
```

Outputs are written to `Output Data Repository/Case Studies/{case_study_name}/`.

---

## 5. Manual_Input.py Reference

All pipeline settings live in `Manual_Input.py`. The file is divided into seven sections ordered by how often each needs changing.

### Section 1 — Run Mode

```python
run_mode = 'Generic'   # or 'Specific'
```

Controls which downstream settings are active. Set this first.

### Section 2 — Year Range

```python
start_year     = 2017
end_year_limit = 2026
```

The pipeline processes financial years from April `start_year` through March `end_year_limit + 1`. Distribution price data must be present for every year in the range (see [Input Data Reference](#6-input-data-reference)).

### Section 3 — Specific Mode Settings

Only used when `run_mode = 'Specific'`.

| Parameter | Type | Description |
|-----------|------|-------------|
| `case_study_name` | str | Subfolder name under `Case Studies/`. Use forward slashes for hierarchy, e.g. `'GroupName/RunName'` |
| `specific_regions` | list of str | DNO region filenames — must match filenames in `Input/Network Prices/Distribution/{year}/` |
| `demand_profile_file` | str | Filename in `Input/Case Study Demand Profiles/`, or `'Flat'` for constant demand |
| `vol_type` | str | `'Non-Domestic Aggregated'`, `'LV'`, `'LV Sub'`, or `'HV'` |
| `has_mic` | bool | `True` if the site has a Metering Import Capacity / Agreed Supply Capacity |
| `residual_band` | str | `'Band 1'` through `'Band 4'` — TCR residual charging band |
| `basis` | float | kW — flat demand level when `demand_profile_file = 'Flat'`; also used as the TRIAD demand proxy for TNUoS estimation when a real demand profile is supplied |
| `power_factor` | float | Site power factor (0–1), used to derive capacity (kVA) and reactive power (kVAr) |

**Voltage type and MIC combinations:**

| `vol_type` | `has_mic` | Connection description |
|------------|-----------|------------------------|
| `'Non-Domestic Aggregated'` | `False` | LV connected, no MIC |
| `'LV'` | `True` | LV connected, with MIC/ASC |
| `'LV Sub'` | `True` | LV Sub connected, with MIC |
| `'HV'` | `True` | HV connected, with MIC/ASC |

### Section 4 — Generic Mode Representative Flat Loads

`generic_flat_loads` is a nested dictionary keyed by `vol_type` then `residual_band`, giving the representative flat demand level in kW for each combination. Defaults are derived from the midpoint of annual consumption ranges per band; Band 4 uses an engineering upper ceiling.

### Section 5 — Build Price Profiles Scope Control

```python
build_scope_mode = 'Full'   # or 'Test'
```

`'Full'` runs the complete Generic scope (all regions, voltage types, bands). `'Test'` restricts the run to the lists in `build_test_vol_types`, `build_test_dno_regions`, and `build_test_bands`, which is useful for faster iteration during development.

### Section 6 — Output Options

```python
apply_vat = False   # True to add a VAT-inclusive column to output workbooks
```

When `True`, a `Total Cost incl. VAT [£]` column (at the standard rate of 20%) is appended to detailed price breakdown outputs alongside the existing ex-VAT figures. Ex-VAT figures are standard for structural cost comparisons, as VAT is typically recoverable by VAT-registered non-domestic consumers.

### Section 7 — Framework Reference Lists

`all_dno_regions` and `all_vol_types` list the fixed set of valid region filenames and voltage types. These should only need updating if the underlying input data naming conventions change.

---

## 6. Input Data Reference

All input data lives under `Code Run Folder/Input/`. The pipeline does not fetch data from the internet — all files must be present before a run.

### Distribution prices

```
Input/Network Prices/Distribution/{year}/{DNO region}.xlsx
```

One workbook per DNO region per financial year (14 regions × years in range). Each workbook must contain an `Annex 1` sheet with the DUoS tariff table. The sheet name varies by year:

- FY2017–FY2020: `Annex 1 LV and HV charges`
- FY2021 onwards: `Annex 1 LV, HV and UMS charges`

Source: DNO Statement of Use of System Charging Methodology (CDCM) documents, published annually by each distribution network operator.

### TNUoS rates

```
Input/Network Prices/Transmission/TNUoS Rates.xlsx
```

A compiled workbook with two sheets:
- **Locational rates**: DNO region × year table of £/kW locational tariff rates
- **Non-locational (residual) rates**: Band × year table of £/consumer/day residual rates (FY2023 onwards, following the TCR reform)

Source: National Grid ESO TNUoS Tariff Report Tables, compiled into the workbook from the supporting documents in `Input/Network Prices/Transmission/Supporting Documents/`.

### BSUoS rates

```
Input/Network Prices/Balancing/BSUoS Rates.xlsx
```

Seasonal BSUoS rates in £/MWh, with a two-row header and columns: Year, Season, BSUoS Tariff. Season values are `winter` and `summer`.

Source: National Grid ESO published BSUoS forecasts and outturn rates.

### Policy levy rates

```
Input/Policy Levy Prices/Policy Levy Rates.xlsx
```

Quarterly policy levy component rates in £/MWh, with columns: Year, Season, Quarter, and one column per levy component (RO, CFD, CM, WHD, and others), plus a `CM` column for the Capacity Market charge used in peak-window calculations.

Quarter convention: Q1 = April–June, Q2 = July–September, Q3 = October–December, Q4 = January–March.

### Wholesale commodity prices

```
Input/Commodity Prices/elexon_MID_{year}.csv
Input/Commodity Prices/compiled_elexon_mid_prices.xlsx
```

Half-hourly settlement period imbalance prices from ELEXON Market Index Data (MID). One CSV per calendar year, plus a compiled workbook. The MID Request Jupyter notebook (`MID Request.ipynb`) in the same folder documents how to download updated data from the ELEXON API.

### Miscellaneous charges

```
Input/Miscellaneous Prices/Miscellaneous Rates.xlsx
```

Annual rates for charges such as metering and data services, in £/year or £/day.

### Demand profile workbooks (Specific mode)

```
Input/Case Study Demand Profiles/{profile name}.xlsx
```

Half-hourly demand data for a real or representative site. The workbook must follow the proforma format documented in `Demand Profile Proforma.xlsx`:

- Header rows in rows 1–4 (rows 0–3, zero-indexed)
- Column names at row 5 (row 4, zero-indexed), including `Date` and `Overall`
- `Date` column: half-hourly datetime values (one row per 30-minute period)
- `Overall` column: half-hourly demand in kW

The profile should span at least one full financial year (April to March). The pipeline uses a financial-year filter internally, so calendar-year profiles that cover April–March are handled correctly.

---

## 7. Output Structure

All outputs are written to `Output Data Repository/`, which is created automatically on the first run.

### Generic mode outputs

```
Output Data Repository/Generic Results/{vol_type}/{band}/
  1 - Price Profiles/Flat Load Analysis/
  2 - Financial Analysis/Flat Load Analysis/
    Summary Workbook.xlsx
    Charts/
```

Each `Summary Workbook.xlsx` contains:

| Sheet | Contents |
|-------|----------|
| `price_type_breakdown_detailed` | Annual cost breakdown with all sub-components, one row per region–year |
| `price_type_breakdown_compressed` | Annual cost breakdown in seven non-overlapping categories, one row per region–year |
| `quarterly_cost_breakdown` | Quarterly cost breakdown, one row per region–year–quarter |
| `cost_per_kwh_detailed` | Detailed costs expressed in p/kWh |
| `cost_per_kwh_compressed` | Compressed costs in p/kWh |
| `key_metrics` | Annual consumption, peak demand, and load factor statistics |

### Specific mode outputs

```
Output Data Repository/Case Studies/{case_study_name}/
  Financial Analysis/Demand Profile Analysis/
    Summary Workbook.xlsx
    Charts/
```

The Specific mode `Summary Workbook.xlsx` has the same sheets as Generic mode, plus additional sheets from the case study analytics module (DUoS band breakdown, TRIAD exposure, ASC headroom).

### Reading output workbooks

The `price_type_breakdown_compressed` sheet is the most reliable source for total cost comparisons. It sums costs into seven non-overlapping categories (DUoS Total, Wholesale, Policy Levy, Miscellaneous, BSUoS, Capacity Market, TNUoS Total) with no double-counting. The `detailed` sheet provides sub-component granularity but includes both aggregate and leaf columns — use the compressed sheet when you need a single total figure.

---

## 8. Extending to New Financial Years

To extend the pipeline to cover an additional financial year:

1. **Distribution prices**: add one workbook per DNO region to `Input/Network Prices/Distribution/{new_year}/` (14 files). Filenames must match the existing naming convention (`10 - Eastern.xlsx` through `23 - Yorkshire.xlsx`).

2. **TNUoS rates**: add the new year's column to both sheets in `Input/Network Prices/Transmission/TNUoS Rates.xlsx`. Source data is in `Supporting Documents/`.

3. **BSUoS rates**: add winter and summer rows for the new financial year to `Input/Network Prices/Balancing/BSUoS Rates.xlsx`.

4. **Policy levy rates**: add four quarterly rows (Q1–Q4) for the new year to `Input/Policy Levy Prices/Policy Levy Rates.xlsx`.

5. **Wholesale prices**: add `elexon_MID_{new_year}.csv` (calendar year) to `Input/Commodity Prices/` and update `compiled_elexon_mid_prices.xlsx`.

6. **Update Manual_Input.py**: set `end_year_limit` to the new financial year.

---

## 9. Case Study Supporting Files

The `Case Study Supporting Files/` subfolder contains reference documents and helper scripts for the case study workflow. Neither file is part of the core pipeline and neither is imported by the main scripts.

### Case_Study_Parameters.py

A reference document (not an executable script) listing all run parameters for the two case studies:

- **ASC Reduction Sweep** — seven runs varying the Agreed Supply Capacity from 11,500 to 17,500 kVA in 1,000 kVA steps, holding the demand profile constant
- **TOU Demand Shifting** — thirteen runs: one unshifted baseline and twelve scenarios shifting Red-band demand forward by 1 hour or backward by 1–3 hours at three magnitudes (10%, 20%, 30%)

To reproduce a case study run, copy the relevant values from this file into `Manual_Input.py` and run `Post Analysis.py`.

### Generate_TOU_Shift_Profiles.py

Generates the 12 shifted demand profile workbooks required by the TOU Demand Shifting case study. Run this script once before running those scenarios:

```bash
python "Case Study Supporting Files/Generate_TOU_Shift_Profiles.py"
```

**Configuration** — edit the two variables at the top of the script:

```python
BASELINE_FILENAME = 'Imperial College London 2016 Demand Profile.xlsx'
SCENARIO_PREFIX   = 'ICL 2016'
```

`BASELINE_FILENAME` must be present in `Input/Case Study Demand Profiles/`. The script writes 12 output files to the same folder, named `{SCENARIO_PREFIX} TOU Shift {direction} {magnitude}.xlsx`.

**Shift mechanic:** for each half-hour slot T in the Red band (16:00–18:30), the script relocates `SHIFT_PCT × original_demand[T]` to slot `T + shift_hours`. Excesses are calculated from the unmodified original values to avoid cascading effects. Total daily kWh is conserved by construction; any slot whose destination falls outside the day has its excess returned to the source slot.
