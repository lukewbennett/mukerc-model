#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Post Analysis v8
================
Script 4 of 4 in the MUKERC pipeline.

Reads half-hourly price profiles produced by Build Price Profiles v8.py and
produces financial analysis outputs for every voltage type, residual charging
band, and DNO region.

Outputs include:
  - Annual price-type cost breakdowns (detailed and compressed)
  - Peak vs. off-peak and seasonal cost splits
  - TNUoS (locational + residual) cost estimates
  - Template-day average price charts (Generic mode)
  - Regional comparison bar charts and annual pie charts

Run behaviour is controlled by Manual_Input.py:
  - run_mode           : 'Generic' (full sweep) or 'Specific' (named case study)
  - start_year         : first financial year to include
  - end_year_limit     : last financial year to include
  - demand_profile_file: 'Flat' or path to a demand profile workbook (Specific only)
  - vol_type / residual_band: connection type and charging band (Specific only)

Prerequisite: Build Price Profiles.py must have been run first.

Created on Sun Jun 07 2026
@author: Luke Bennett
"""

import os
import sys
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import matplotlib as mpl
import datetime
from datetime import time
import re
import time as pytime

plt.style.use('seaborn-v0_8-whitegrid')

# --- START TIMER ---
start_time = pytime.perf_counter()

# =============================================================================
# SECTION 1: INPUT PARAMETERS
# =============================================================================

from Manual_Input import (
    run_mode, case_study_name, specific_regions, all_dno_regions,
    all_vol_types, vol_type, has_mic, residual_band,
    basis, capacity, reactive_power, power_factor, plot_pk_start, plot_pk_end,
    demand_profile_file, generic_flat_loads,
    start_year, end_year_limit, apply_vat,
    build_scope_mode, build_test_vol_types,
    build_test_dno_regions, build_test_bands
)

from Modules.Helper_Functions import (
    get_vol_row_label, get_annex_sheet,
    prepare_tariff_tables, prepare_seasonal_component_tables,
    load_bsuos_rates
)
from Modules.Case_Study_Analytics import (
    duos_band_breakdown, blended_duos_rate,
    triad_exposure_share, exceeded_cap_incidence, headroom_utilisation
)

# NOTE: price_type is hardcoded as 'Import' throughout this script.
# Export functionality has been removed from the MUKERC framework.
price_type = 'Import'

# =============================================================================
# DEMAND PROFILE HELPER FUNCTIONS
# Defined here (before the validation block) so they can be called at startup.
# =============================================================================

def _find_demand_header_row(profile_path):
    """
    Scans column A (up to 20 rows) for the cell containing 'Date'.
    Returns that row index so pd.read_excel can skip metadata rows above it.
    Falls back to 0 if no metadata rows are present (old-style files).
    """
    probe = pd.read_excel(profile_path, header=None, usecols=[0], nrows=20)
    for idx, val in probe.iloc[:, 0].items():
        if str(val).strip().lower() == 'date':
            return int(idx)
    return 0


def _detect_profile_stats(profile_path):
    """
    Reads the demand profile and returns consumption statistics used to suggest
    the correct vol_type for the site.

    Voltage type heuristic (peak kW — physical confirmation always recommended):
      < 10 kW mean  → Non-Domestic Aggregated
      peak < 300 kW → LV
      peak < 2,500  → LV Sub
      peak ≥ 2,500  → HV

    Note: residual band thresholds differ by voltage type and are not checked
    here — the user must confirm the correct band in Manual_Input.py.
    """
    header_row = _find_demand_header_row(profile_path)
    df = pd.read_excel(profile_path, header=header_row, usecols=['Date', 'Overall'])
    df['Overall'] = pd.to_numeric(df['Overall'], errors='coerce')
    df['Date']    = pd.to_datetime(df['Date'],    errors='coerce')
    df = df.dropna(subset=['Date', 'Overall'])

    annual_mwh = df['Overall'].sum() * 0.5 / 1000   # kW × 0.5 h = kWh ÷ 1000 = MWh
    peak_kw    = df['Overall'].max()
    mean_kw    = df['Overall'].mean()

    if mean_kw < 10:
        suggested_vol = 'Non-Domestic Aggregated'
    elif peak_kw < 300:
        suggested_vol = 'LV'
    elif peak_kw < 2_500:
        suggested_vol = 'LV Sub'
    else:
        suggested_vol = 'HV'

    return annual_mwh, peak_kw, mean_kw, suggested_vol


# =============================================================================
# VALIDATION BLOCK — Specific mode only
# =============================================================================

valid_combinations = {
    'Non-Domestic Aggregated': False,
    'LV':                      True,
    'LV Sub':                  True,
    'HV':                      True,
}

if run_mode == 'Specific':

    # ── vol_type / has_mic ────────────────────────────────────────────────────
    if vol_type not in valid_combinations:
        raise ValueError(
            f"Invalid vol_type: '{vol_type}'. "
            f"Must be one of: {list(valid_combinations.keys())}"
        )
    if has_mic != valid_combinations[vol_type]:
        if vol_type == 'Non-Domestic Aggregated':
            raise ValueError(
                f"vol_type = '{vol_type}' requires has_mic = False."
            )
        else:
            raise ValueError(
                f"vol_type = '{vol_type}' requires has_mic = True."
            )

    # ── case_study_name ───────────────────────────────────────────────────────
    if not case_study_name.strip():
        raise ValueError("case_study_name cannot be empty in Manual_Input.py.")

    _cs_output_path = (
        f"../Output Data Repository/Case Studies/{case_study_name}"
    )
    if os.path.exists(_cs_output_path):
        if sys.stdin.isatty():
            try:
                _resp = input(
                    f"\n  ⚠  Case study '{case_study_name}' already exists.\n"
                    f"     Outputs will be merged / overwritten.  Continue? [y/N]: "
                )
            except EOFError:
                _resp = 'y'
            if _resp.strip().lower() != 'y':
                raise SystemExit("Run cancelled — case study not overwritten.")
        else:
            print(f"  ⚠  Case study '{case_study_name}' already exists — overwriting (non-interactive).")

    # ── specific_regions ──────────────────────────────────────────────────────
    _invalid_regions = [r for r in specific_regions if r not in all_dno_regions]
    if _invalid_regions:
        raise ValueError(
            f"Unrecognised region(s) in specific_regions: {_invalid_regions}.\n"
            f"Valid options are: {all_dno_regions}"
        )

    # ── price profile existence + year range ──────────────────────────────────
    _profile_root = (
        f"../Output Data Repository/Generic Results"
        f"/{vol_type}/{residual_band}/1 - Price Profiles/Half-Hourly Prices"
    )
    _missing_parquets = []
    for _reg in specific_regions:
        _pq = f"{_profile_root}/{_reg[:-5]}.parquet"
        if not os.path.exists(_pq):
            _missing_parquets.append(_reg[:-5])
    if _missing_parquets:
        raise FileNotFoundError(
            f"Generic price profiles not found for: {_missing_parquets}\n"
            f"  vol_type='{vol_type}', band='{residual_band}'.\n"
            f"  Run Build Price Profiles v8.py with these settings first."
        )

    _sample_pq = (
        f"{_profile_root}/{specific_regions[0][:-5]}.parquet"
    )
    _pq_years = sorted(
        int(c.split()[-1])
        for c in pd.read_parquet(_sample_pq).columns
        if c.startswith('Year ')
    )
    _missing_years = [
        y for y in range(start_year, end_year_limit + 1)
        if y not in _pq_years
    ]
    if _missing_years:
        raise ValueError(
            f"Generic price profiles cover years {_pq_years[0]}–{_pq_years[-1]} "
            f"but start_year={start_year} / end_year_limit={end_year_limit} "
            f"requests year(s) {_missing_years}.\n"
            f"  Re-run Build Price Profiles v8.py to include these years."
        )

    # ── demand profile: auto-detect band and vol_type ─────────────────────────
    if demand_profile_file != 'Flat':
        _dp_path = (
            f"Input/Case Study Demand Profiles/{demand_profile_file}"
        )
        if not os.path.exists(_dp_path):
            raise FileNotFoundError(
                f"Demand profile not found: '{demand_profile_file}'.\n"
                f"  Expected location: '{_dp_path}'."
            )
        _ann_mwh, _peak_kw, _mean_kw, _sug_vol = (
            _detect_profile_stats(_dp_path)
        )
        print(
            f"\n  ── Demand Profile: '{demand_profile_file}'\n"
            f"     Annual consumption : {_ann_mwh:>10,.0f} MWh\n"
            f"     Peak demand        : {_peak_kw:>10,.0f} kW"
            f"   →  suggested voltage: {_sug_vol}\n"
            f"     Mean demand        : {_mean_kw:>10,.0f} kW\n"
            f"     Active settings    :  {vol_type}  |  {residual_band}"
        )
        if _sug_vol != vol_type:
            if sys.stdin.isatty():
                _resp = input(
                    f"\n  ⚠  Demand level suggests '{_sug_vol}' but "
                    f"vol_type='{vol_type}'.  Continue anyway? [y/N]: "
                )
                if _resp.strip().lower() != 'y':
                    raise SystemExit(
                        "Run cancelled — update vol_type in Manual_Input.py."
                    )
            else:
                print(
                    f"  ⚠  Demand suggests '{_sug_vol}' but vol_type='{vol_type}' "
                    f"— proceeding (non-interactive)."
                )

    print(f"\n✓ Validation passed: vol_type='{vol_type}', has_mic={has_mic}, "
          f"band='{residual_band}'")

# =============================================================================
# DETERMINE RUN CONFIGURATION FROM MODE
# =============================================================================

if run_mode == 'Generic':
    if build_scope_mode == 'Test':
        run_vol_types   = build_test_vol_types
        run_dno_regions = build_test_dno_regions
        bands_to_run    = build_test_bands
    else:
        run_vol_types   = all_vol_types
        run_dno_regions = all_dno_regions
        bands_to_run    = ['Band 1', 'Band 2', 'Band 3', 'Band 4']
    output_root  = "../Output Data Repository/Generic Results"
    generic_root = output_root
    print(f"\n✓ Run Mode: Generic ({build_scope_mode} scope)")
else:
    # Stage 6: Specific mode reads price profiles from Generic Results
    # and writes only financial outputs to Case Studies
    run_vol_types    = [vol_type]
    run_dno_regions  = specific_regions
    output_root      = (
        f"../Output Data Repository/Case Studies/{case_study_name}"
    )
    generic_root     = "../Output Data Repository/Generic Results"
    bands_to_run     = [residual_band]
    print(f"\n✓ Run Mode: Specific — '{case_study_name}'")
    print(
        f"  Price profiles read from: Generic Results/"
        f"{vol_type}/{residual_band}"
    )
    print(f"  Financial outputs written to: Case Studies/{case_study_name}")

# =============================================================================
# ACTIVE ANALYSIS BASIS
# =============================================================================

if run_mode == 'Generic':
    analysis_basis     = 'flat_load'
    analysis_label     = 'Flat Load Analysis'
    use_flat_load      = True
    use_demand_profile = False
else:
    if demand_profile_file == 'Flat':
        analysis_basis     = 'flat_load'
        analysis_label     = 'Flat Load Analysis'
        use_flat_load      = True
        use_demand_profile = False
    else:
        analysis_basis     = 'demand_profile'
        analysis_label     = 'Demand Profile Analysis'
        use_flat_load      = False
        use_demand_profile = True

if apply_vat:
    analysis_label = f"{analysis_label} (incl. VAT)"

print(f"✓ Active analysis basis: {analysis_label}")

# =============================================================================
# OUTPUT FOLDER LABELS
# Generic mode retains numbered folders for structured bulk outputs.
# Specific mode uses flatter, unnumbered folders for one-case usability.
# =============================================================================

if run_mode == 'Generic':
    financial_dir_name = '2 - Financial Analysis'  # Financial results root for bulk Generic outputs
    reference_dir_name = 'Reference Data'  # Shared reference-data root for bulk Generic outputs
    chart_dir_name = '3 - Volumetric Price Profiles'  # Chart output root for bulk Generic outputs
else:
    financial_dir_name = 'Financial Analysis'  # Financial results root for a single case-study output
    reference_dir_name = 'Reference Data'  # Reference-data root for a single case-study output
    chart_dir_name = 'Volumetric Price Profiles'  # Chart output root for a single case-study output

# =============================================================================
# SHARED SETUP
# =============================================================================

def base_calc(tme, delta):
    """Returns a datetime.time shifted by delta from tme (used to derive chart peak window edges)."""
    return (
        datetime.datetime.combine(datetime.date.today(), tme) + delta
    ).time()

base_start = base_calc(plot_pk_end,   datetime.timedelta(minutes=30))
base_end   = base_calc(plot_pk_start, datetime.timedelta(minutes=-30))
peak_label = (
    str(plot_pk_start)[:5] + " - " + str(base_start)[:5]
).replace(':', ',')

years_past_all = sorted([
    f for f in os.listdir("Input/Network Prices/Distribution")
    if f != ".DS_Store"
])
years_past = [
    yr for yr in years_past_all
    if start_year <= int(yr) <= end_year_limit
]

if not years_past:
    raise ValueError(
        f"No distribution input years found between "
        f"start_year={start_year} and end_year_limit={end_year_limit}."
    )

latest_run_year = years_past[-1]

# years_in is identical for every voltage type and band — derived once from
# years_past (the set of years with distribution input data in scope).
years_in = pd.Index(years_past)

# template_dates_full maps every half-hour to its month, weekday, time, and
# week-in-month. It depends only on the year range, so it is built once here.
template_dates_full = pd.DataFrame()
for _yr in list(years_in.astype(int)):
    _yr1 = _yr + 1
    _date_range = pd.date_range(
        start=f'{_yr}-04-01', end=f'{_yr1}-04-01', freq='30min'
    )[:-1]
    _td = pd.DataFrame({f'Year {_yr}': _date_range})
    _td[f'Month {_yr}']         = _td[f'Year {_yr}'].dt.month
    _td[f'Day in Week {_yr}']   = _td[f'Year {_yr}'].dt.weekday
    _td[f'Time {_yr}']          = _td[f'Year {_yr}'].dt.time
    _td[f'Week in Month {_yr}'] = (_td[f'Year {_yr}'].dt.day - 1) // 7 + 1
    _td[f'Date Indexes {_yr}']  = list(zip(
        _td[f'Month {_yr}'], _td[f'Day in Week {_yr}'],
        _td[f'Week in Month {_yr}'], _td[f'Time {_yr}']
    ))
    template_dates_full = pd.concat([template_dates_full, _td], axis=1)

tnuos_locational_df = pd.read_excel(
    "Input/Network Prices/Transmission/TNUoS Rates.xlsx",
    sheet_name='HH Locational Demand Tariff',
    header=3
)
tnuos_nonlocational_df = pd.read_excel(
    "Input/Network Prices/Transmission/TNUoS Rates.xlsx",
    sheet_name='Non-Locational Banded Charges',
    header=3
)

tariff_sum = pd.read_excel(
    "Input/Policy Levy Prices/Policy Levy Rates.xlsx",
    header=3
)  # Policy Levy Rates workbook containing Policy Levy component tariffs and capacity market rate

misc_charges = pd.read_excel(
    "Input/Miscellaneous Prices/Miscellaneous Rates.xlsx",
    header=3
)  # Seasonal miscellaneous charge workbook with the same year/season structure as the Policy Levy rates workbook

bsuos_rates = load_bsuos_rates(
    "Input/Network Prices/Balancing/BSUoS Rates.xlsx"
)  # Seasonal BSUoS rates in £/MWh; divided by 10 at point of use to convert to p/kWh

pl_tariff_cols, pl_tariff_table, pl_tariff_total, cm_table = (
    prepare_tariff_tables(tariff_sum)
)  # Dynamic Policy Levy tariff interpretation shared with Build Profiles

misc_charge_cols, misc_charge_table, misc_charge_total = (
    prepare_seasonal_component_tables(misc_charges)
)  # Dynamic miscellaneous seasonal charge interpretation shared with Build Profiles

# =============================================================================
# HELPER FUNCTIONS
# =============================================================================

def get_tnuos_nonloc_label(current_vol_type, band_number):
    """Returns the TNUoS non-locational band label used to look up residual charges in the Annex input."""
    if current_vol_type == 'Non-Domestic Aggregated':
        return f'LV_NoMIC_{band_number}'
    elif current_vol_type in ['LV', 'LV Sub']:
        return f'LV{band_number}'
    elif current_vol_type == 'HV':
        return f'HV{band_number}'


def read_tnuos_locational(dno_no, year, df):
    """Reads TNUoS HH locational tariff [£/kW] from pre-loaded DataFrame."""
    return float(df[df['DNO Region'] == dno_no][int(year)].iat[0])


def read_tnuos_nonlocational(current_vol_type, band_number, year, df):
    """Reads TNUoS non-locational band charge [£/day] from pre-loaded DataFrame."""
    band_label = get_tnuos_nonloc_label(current_vol_type, band_number)
    return float(df[df['Band'] == band_label][int(year)].iat[0])


def split_data(labels, text, time_type):
    """
    Scans descriptive LLF period text and returns the start/end month or
    day integer bounds implied by the first two matching label strings.

    Parameters
    ----------
    labels    : list   ordered label strings to search for (e.g. months or days)
    text      : str    raw descriptive text from the Annex 5 table cell
    time_type : str    'Month' or 'Day' — controls the 'all year' fallback value
    """
    count = mon_1 = mon_2 = 0
    mon1_str = mon2_str = ""
    for idx, item in enumerate(labels):
        if item in text:
            count += 1
            if count == 1:
                mon1_str, mon_1 = item, idx + 1
            elif count == 2:
                mon2_str, mon_2 = item, idx + 1
    if mon_1 > 0 and mon_2 > 0:
        if text.find(mon1_str) > text.find(mon2_str):
            mon_1, mon_2 = mon_2, mon_1
    elif mon_1 == 0 and mon_2 == 0:
        if 'all year' in text:
            mon_1 = 1
            mon_2 = 12 if time_type == 'Month' else 7
    elif mon_2 == 0:
        mon_2 = mon_1
    return mon_1, mon_2


def peak_base_calc(data, dno_region, yr, peak_base_cols):
    """Calculates peak and off-peak annual costs from half-hourly data."""
    peak_time = data.between_time(plot_pk_start, plot_pk_end)
    base_time = data.between_time(base_start, base_end)
    peak      = peak_time['Overall'].sum()
    n_peak    = base_time['Overall'].sum()
    peak_base = pd.DataFrame(
        data=[yr, peak, n_peak],
        index=peak_base_cols,
        columns=[dno_region]
    ).T
    return peak_base.set_index('Year', append=True)


def sum_winter_calc(data, dno_region, yr, win_sum_cols):
    """Calculates winter and summer annual costs from half-hourly data."""
    stacked = data[['Overall']]
    winter  = stacked.loc[
        stacked.index.month.isin([10, 11, 12, 1, 2, 3])
    ].resample('YS').sum().sum().iat[0]
    summer  = stacked.loc[
        stacked.index.month.isin([4, 5, 6, 7, 8, 9])
    ].resample('YS').sum().sum().iat[0]
    result  = pd.DataFrame(
        data=[yr, winter, summer],
        index=win_sum_cols,
        columns=[dno_region]
    ).T
    return result.set_index('Year', append=True)


def llf(current_vol_type, vol_llf, dno_list, years_past_list, root_out, save_reference=True):
    """
    Reads DNO Annex 5 LLF sheets for every region and year, building two outputs:

    llfprices_full : DataFrame indexed by (Year, DNO Region) containing
                     Period 1..5 LLF multipliers for the given voltage level.
                     Period 1 is used for the TNUoS locational cost calculation.
    dftime_tup_out : DataFrame indexed by (Year, Category, DNO Region) containing
                     the time-band definitions associated with each LLF period.

    Both are also saved as reference Excel files under root_out/…/LLFs/.
    Regions or years for which the Annex 5 sheet is absent are skipped silently.
    """
    days_list   = ['mon','tue','wed','thu','fri','sat','sun']
    months_list = ['jan','feb','mar','apr','may','jun',
                   'jul','aug','sep','oct','nov','dec']
    key_vars    = pd.Series([
        'Day Start','Day End','Month Start','Month End',
        'Time 1','Time 2','Time 3','Time 4','Time 5','Time 6'
    ])
    llf_times = key_vars[4:]

    llfprices_list = []
    dftime_list    = []

    for region in list(dno_list):
        dno_region = region[:-5]
        years_rev  = years_past_list[::-1]

        for llf_yr in list(years_rev):
            time_periods = [
                'Period 1','Period 2','Period 3','Period 4','Period 5'
            ]
            llf_direct = os.path.abspath(os.path.dirname(
                f"Input/Network Prices/Distribution/{llf_yr}/{region}"
            ))

            try:
                l = pd.read_excel(
                    f"{llf_direct}/{region}",
                    sheet_name='Annex 5 LLFs'
                )
            except FileNotFoundError:
                continue

            try:
                lett_list = [
                    'A','B','C','D','E','F','G','H','I','J','K','L','M'
                ]
                l.columns = lett_list[:len(l.T)]
                l['A']    = (
                    l['A'].str.replace('-',' ').str.lower().str.strip()
                )
                l1         = l.set_index(l.columns[0])
                times_tab  = l1.copy()
                times_tab.columns  = l1.loc['time periods'].values
                prices_tab = l1.copy()
                prices_tab.columns = l1.loc['metered voltage'].values

                try:
                    table_times  = times_tab[time_periods]
                    table_prices = prices_tab[time_periods]
                except KeyError:
                    time_periods = time_periods[:-1]
                    table_times  = times_tab[time_periods]
                    table_prices = prices_tab[time_periods]

                llf_prices   = table_prices.loc[vol_llf]
                llf_pricesdf = pd.DataFrame(llf_prices).T
                llf_pricesdf['Year']       = str(llf_yr)
                llf_pricesdf['DNO Region'] = dno_region

                if llf_prices.loc['Period 1'] > 0:
                    for category in time_periods:
                        select  = (
                            table_times[[category]]
                            if isinstance(table_times[category], pd.Series)
                            else table_times[category]
                        )
                        select1 = select.loc[
                            :, ~select.columns.duplicated()
                        ]
                        select2 = select1[
                            select1[category].str.contains(
                                '00|30|All', na=False
                            )
                        ].reset_index()
                        select3 = select2[
                            select2[category].str.len() < 50
                        ].copy()
                        select3['A']      = select3['A'].str.lower()
                        select3[category] = select3[category].str.lower()
                        select4 = select3[
                            ~select3.A.str.contains("note")
                        ]

                        for i in range(len(select4)):
                            text_data = select4.iloc[i]
                            raw_text  = str(text_data['A'])

                            month_start, month_end = split_data(
                                months_list, raw_text, 'Month'
                            )
                            day_start, day_end = split_data(
                                days_list, raw_text, 'Day'
                            )
                            ind_time = str(text_data[category])

                            row_dict = {
                                'Year':        str(llf_yr),
                                'Category':    category,
                                'DNO Region':  dno_region,
                                'Day Start':   day_start,
                                'Day End':     day_end,
                                'Month Start': month_start,
                                'Month End':   month_end
                            }

                            if 'all' in ind_time:
                                row_dict['Time 1'] = ind_time
                            else:
                                ind_time = re.sub(
                                    "[^0-9]", "", ind_time
                                ).replace('2400', '2330')
                                for idx, t in enumerate(llf_times):
                                    chunk = ind_time[idx*4 : (idx+1)*4]
                                    row_dict[t] = (
                                        chunk if chunk != '' else np.nan
                                    )

                            if pd.notna(row_dict.get('Time 1')):
                                dftime_list.append(row_dict)

                    print(f"  LLF Processed: {dno_region} — {llf_yr}")
                    llfprices_list.append(llf_pricesdf)

            except Exception as e:
                print(
                    f"  Warning: Skipped LLF {dno_region} ({llf_yr}): {e}"
                )

    llfprices_full = (
        pd.concat(llfprices_list) if llfprices_list else pd.DataFrame()
    )
    if not llfprices_full.empty:
        llfprices_full.set_index(['Year', 'DNO Region'], inplace=True)

    dftime_tup_out = pd.DataFrame(dftime_list)

    if not dftime_tup_out.empty:
        dftime_tup_out.set_index(
            ['Year', 'Category', 'DNO Region'], inplace=True
        )

        t1_shifted = dftime_tup_out['Time 1'].shift()
        t2_shifted = dftime_tup_out['Time 2'].shift()
        mask = dftime_tup_out['Time 1'] == 'all other times'
        dftime_tup_out.loc[mask, 'Time 1'] = t2_shifted[mask]
        dftime_tup_out.loc[mask, 'Time 2'] = t1_shifted[mask]

        dftime_tup_out['Time 1'] = dftime_tup_out['Time 1'].str.replace(
            '2330|0030', '0000', regex=True
        )
        dftime_tup_out['Time 2'] = dftime_tup_out['Time 2'].str.replace(
            '0030|0000', '2330', regex=True
        )

        if save_reference:
            llf_save_dir = f"{root_out}/Reference Data/LLFs"
            os.makedirs(llf_save_dir, exist_ok=True)
            dftime_tup_out.to_excel(
                f"{llf_save_dir}/Time Periods All Regions.xlsx"
            )
            llfprices_full.to_excel(
                f"{llf_save_dir}/LLF Factors All Regions.xlsx"
            )

    return llfprices_full, dftime_tup_out


def prep_costs(df):
    """Preserves explicit DUoS component columns and calculates grand total.

    FIXED 14/08/2026 (Phase 8 audit): previously did out['Total Cost [£]'] =
    out.sum(axis=1), a blind row-sum across every column in df. In the
    use_demand_profile (Specific mode, real demand profile) path, df already
    carries: (a) a pre-existing, PARTIAL 'Total Cost [£]' column (set upstream
    from only 6 of the ~13 true cost categories, missing DUoS Fixed/Capacity/
    Reactive/Exceeded and TNUoS entirely), and (b) aggregate columns
    ('Policy Levy Costs Cost [£]', 'Est. TNUoS Total Cost [£]') sitting
    alongside the individual leaf columns those aggregates are themselves
    sums of. A blind row-sum therefore triple-counted some components and
    self-included the stale partial total, inflating Total Cost by roughly
    2.2x for both Chapter 5 case studies (confirmed directly: £250.33m vs the
    correct £112.89m for the ASC/TOU 16,500 kVA 9-year reference, verified by
    independently summing the same non-overlapping categories
    build_compressed_price_breakdown() uses). Generic mode (flat-load,
    Chapter 4) was not affected, since its cost table never carries a
    pre-existing 'Total Cost [£]' column at this point.
    Fix: compute Total Cost from the same well-defined, non-overlapping
    top-level categories build_compressed_price_breakdown() already uses
    (and which the compressed price-breakdown sheet already reported
    correctly throughout), rather than summing every column blindly.
    """
    out = df.copy()
    _compressed_for_total = build_compressed_price_breakdown(df)
    out['Total Cost [£]'] = _compressed_for_total.sum(axis=1)  # Grand total annual cost, non-overlapping categories only
    out.index = pd.MultiIndex.from_tuples(
        out.index, names=['DNO Region', 'Year']
    )
    return out.sort_values('Year')


def load_active_demand_profile(profile_path):
    """
    Loads and validates the user-selected demand profile.

    Expected format (Input/Case Study Demand Profiles/):
      - Optional metadata rows at the top (case study name, description,
        unit notes). These are skipped automatically — the function scans
        for the row whose first cell is 'Date' and uses that as the header.
      - Header row: Date | Overall | [optional component columns …]
      - Data rows: half-hourly datetime in 'Date', demand in kW in 'Overall'.
      - The profile may start in any month; it does not need to align with the
        April financial-year start. Cost calculation uses a (Month, DayOfWeek,
        WeekInMonth, Time) template index so the calendar offset is irrelevant.
        TRIAD estimation requires dates in Nov–Feb on weekdays.
    """
    header_row = _find_demand_header_row(profile_path)
    profile_df = pd.read_excel(profile_path, header=header_row)
    profile_df.columns = profile_df.columns.astype(str).str.strip()

    required_cols = ['Date', 'Overall']
    missing_cols  = [
        c for c in required_cols if c not in profile_df.columns
    ]
    if missing_cols:
        raise ValueError(
            f"Demand profile '{os.path.basename(profile_path)}' "
            f"is missing required column(s): {missing_cols}."
        )

    profile_df['Date'] = pd.to_datetime(profile_df['Date'])
    demand_value_cols  = [c for c in profile_df.columns if c != 'Date']
    for col in demand_value_cols:
        profile_df[col] = pd.to_numeric(profile_df[col], errors='coerce')

    if profile_df['Overall'].isna().all():
        raise ValueError(
            f"Demand profile '{os.path.basename(profile_path)}' "
            f"contains no valid numeric values in 'Overall'."
        )

    profile_df['Overall'] = profile_df['Overall'].fillna(
        profile_df['Overall'].mean()
    )

    raw_component_cols = [
        c for c in profile_df.columns
        if c not in ['Date', 'Overall', 'Other']
    ]
    valid_component_cols = []
    for col in raw_component_cols:
        if profile_df[col].notna().any():
            profile_df[col] = profile_df[col].fillna(
                profile_df[col].mean()
            )
            valid_component_cols.append(col)
        else:
            print(
                f"Warning: Dropped empty column '{col}' from "
                f"'{os.path.basename(profile_path)}'."
            )

    if valid_component_cols:
        if 'Other' in profile_df.columns:
            if profile_df['Other'].notna().any():
                profile_df['Other'] = profile_df['Other'].fillna(
                    profile_df['Other'].mean()
                )
            else:
                profile_df['Other'] = 0.0
        else:
            profile_df['Other'] = (
                profile_df['Overall'] -
                profile_df[valid_component_cols].sum(axis=1)
            )
        if (profile_df['Other'] < 0).any():
            print(
                f"Warning: Derived 'Other' contains negative values in "
                f"'{os.path.basename(profile_path)}'."
            )
        active_usage_cols = valid_component_cols + ['Other']
    else:
        active_usage_cols = []

    active_profile_fin        = profile_df.copy()
    active_profile_structured = profile_df.copy()

    active_profile_structured['Month'] = (
        active_profile_structured['Date'].dt.month
    )
    active_profile_structured['Day in Week'] = (
        active_profile_structured['Date'].dt.weekday
    )
    active_profile_structured['Time'] = (
        active_profile_structured['Date'].dt.time
    )
    active_profile_structured['Week in Month'] = (
        active_profile_structured['Date'].dt.day - 1
    ) // 7 + 1
    active_profile_structured['Date Indexes'] = list(zip(
        active_profile_structured['Month'],
        active_profile_structured['Day in Week'],
        active_profile_structured['Week in Month'],
        active_profile_structured['Time']
    ))

    return active_profile_structured, active_profile_fin, active_usage_cols


def create_run_settings_tables(
    run_mode, case_study_name, current_vol_type, current_band,
    analysis_basis, analysis_label, demand_profile_file,
    current_basis, capacity, reactive_power,
    run_dno_regions, years_in, use_flat_load, use_demand_profile
):
    """Builds reference tables documenting the active run configuration."""
    configuration_df = pd.DataFrame({
        'setting': [
            'run_mode', 'case_study_name', 'current_vol_type',
            'current_band', 'analysis_basis', 'analysis_label',
            'demand_profile_file', 'current_basis', 'capacity',
            'reactive_power'
        ],
        'value': [
            run_mode, case_study_name, current_vol_type,
            current_band, analysis_basis, analysis_label,
            demand_profile_file, current_basis, capacity,
            reactive_power
        ]
    })
    regions_df = pd.DataFrame({
        'region_file': run_dno_regions,
        'dno_region':  [r[:-5] for r in run_dno_regions]
    })
    years_df = pd.DataFrame({'year': list(years_in)})
    flags_df = pd.DataFrame({
        'flag':  ['use_flat_load', 'use_demand_profile'],
        'value': [use_flat_load, use_demand_profile]
    })
    return {
        'configuration':  configuration_df,
        'regions':        regions_df,
        'years':          years_df,
        'analysis_flags': flags_df
    }


def create_summary_tables(
    analysis_costs_prepped, peak_split_full, season_split_full,
    use_costs_full, use_demand_profile, partial_yr_set=None
):
    """Builds summary output tables for the active analysis branch.

    partial_yr_set: set of string year keys (e.g. {'2026'}) identified as
    incomplete financial years (fewer than 17520 half-hourly slots), per the
    detection already performed by the caller. Any table that collapses the
    Year dimension via a mean/aggregate (average_annual_costs, key_metrics)
    must exclude these years first, since a part-year total silently pulled
    into a same-weight average understates the true annual figure. Tables
    that retain Year as an explicit column or index level (annual_total_costs,
    the price_type_breakdown tables) are left with every year present, since
    their own downstream consumers already filter by complete financial year
    where relevant (see e.g. chapter5_cross_run_analysis.py's LAST_FULL_YEAR).
    """
    annual_total_costs_df = analysis_costs_prepped[
        ['Total Cost [£]']
    ].reset_index().pivot(
        index='DNO Region', columns='Year', values='Total Cost [£]'
    )

    partial_yr_set = partial_yr_set or set()
    if partial_yr_set:
        analysis_costs_full_years = analysis_costs_prepped[
            ~analysis_costs_prepped.index
            .get_level_values('Year').astype(str)
            .isin(partial_yr_set)
        ]  # Complete-financial-years-only view, used only for the two
           # tables below that average/aggregate across years
    else:
        analysis_costs_full_years = analysis_costs_prepped

    average_annual_costs_df = (
        analysis_costs_full_years
        .groupby(level='DNO Region').mean()
        .sort_values('Total Cost [£]')
    )
    price_type_breakdown_detailed_df = build_detailed_price_breakdown(
        analysis_costs_prepped
    ).reset_index()  # Detailed annual price breakdown table included in the summary workbook for audit and inspection

    price_type_breakdown_compressed_df = build_compressed_price_breakdown(
        analysis_costs_prepped
    ).reset_index()  # Compressed annual price breakdown table included in the summary workbook for readability

    seasonal_split_export = peak_or_season_table_for_export(
        season_split_full
    )  # Seasonal split export table for the summary workbook

    peak_split_export = peak_or_season_table_for_export(
        peak_split_full
    )  # Peak vs off-peak split export table for the summary workbook

    key_metrics_df = build_key_metrics_table(
        analysis_costs_full_years
    )  # Compact summary table highlighting regional cost spread and extrema; excludes partial years for the same reason as average_annual_costs above

    summary_tables = {
        'annual_total_costs': annual_total_costs_df,
        'average_annual_costs': average_annual_costs_df,
        'price_type_breakdown_compressed': price_type_breakdown_compressed_df,
        'price_type_breakdown_detailed': price_type_breakdown_detailed_df,
        'seasonal_split': seasonal_split_export,
        'peak_vs_off_peak': peak_split_export,
        'key_metrics': key_metrics_df
    }

    if use_demand_profile and not use_costs_full.empty:
        summary_tables['demand_type_breakdown'] = (
            use_costs_full.reset_index()
        )

    return summary_tables


def peak_or_season_table_for_export(df):
    """Standardises peak or seasonal split tables for workbook export."""
    export_df = df.copy()
    export_df = export_df.rename(columns={
        c: f'{c} [£]' for c in export_df.columns if '[£]' not in c
    })
    if 'Total Cost [£]' not in export_df.columns:
        export_df['Total Cost [£]'] = export_df.sum(axis=1)
    export_df.index = pd.MultiIndex.from_tuples(
        export_df.index, names=['DNO Region', 'Year']
    )
    return export_df.reset_index()


def build_key_metrics_table(analysis_costs_prepped):
    """Builds a compact key-metrics summary table."""
    avg          = analysis_costs_prepped.groupby(level='DNO Region').mean()
    min_region   = avg['Total Cost [£]'].idxmin()
    max_region   = avg['Total Cost [£]'].idxmax()
    min_cost     = avg.loc[min_region, 'Total Cost [£]']
    max_cost     = avg.loc[max_region, 'Total Cost [£]']
    return pd.DataFrame({
        'metric': [
            'lowest_average_cost_region', 'lowest_average_cost_gbp',
            'highest_average_cost_region', 'highest_average_cost_gbp',
            'average_cost_spread_gbp'
        ],
        'value': [
            min_region, min_cost, max_region, max_cost,
            max_cost - min_cost
        ]
    })


def save_summary_workbook(summary_tables, summary_workbook_path):
    """Saves summary tables to a single workbook."""
    with pd.ExcelWriter(
        summary_workbook_path, engine='openpyxl'
    ) as writer:
        for sheet_name, summary_df in summary_tables.items():
            write_index = not isinstance(
                summary_df.index, pd.RangeIndex
            )
            summary_df.to_excel(
                writer, sheet_name=sheet_name, index=write_index
            )


def reorder_columns_for_export(df, preferred_order):
    """Reorders columns using a preferred export order."""
    ordered_cols   = [c for c in preferred_order if c in df.columns]
    remaining_cols = [c for c in df.columns if c not in ordered_cols]
    return df[ordered_cols + remaining_cols]


def round_export_table(df, decimals=2):
    """Rounds numeric columns in an export table."""
    out          = df.copy()
    numeric_cols = out.select_dtypes(include=[np.number]).columns
    out[numeric_cols] = out[numeric_cols].round(decimals)
    return out

def build_effective_cost_benchmark(analysis_costs_prepped, annual_energy_kwh):
    """Converts annual total costs to an effective electricity rate (p/kWh) for Generic flat-load results."""
    effective_cost_df = analysis_costs_prepped[['Total Cost [£]']].copy()  # Annual total bill table used to derive effective electricity cost

    effective_cost_df['Effective Cost [p/kWh]'] = (
        effective_cost_df['Total Cost [£]'] / annual_energy_kwh
    ) * 100  # All-in annual electricity cost converted from £ total to p/kWh using Generic annual energy demand

    annual_p_per_kwh_df = effective_cost_df.reset_index().pivot(
        index='DNO Region', columns='Year', values='Effective Cost [p/kWh]'
    )  # Regional annual effective electricity cost benchmark in p/kWh

    return annual_p_per_kwh_df

def save_effective_cost_benchmark(annual_p_per_kwh_df, workbook_path):
    """Saves the Generic flat-load effective electricity cost benchmark to a single-sheet workbook."""
    with pd.ExcelWriter(workbook_path, engine='openpyxl') as writer:
        annual_p_per_kwh_df.to_excel(
            writer, sheet_name='annual_p_per_kwh'
        )

def build_detailed_price_breakdown(df):
    """Reorders cost columns into the canonical detailed breakdown order, preserving all explicit DUoS and Policy Levy components."""
    detailed_df = df.copy()

    # In Specific (demand-profile) mode, individual pl_cost_cols are present and the
    # upstream aggregate 'Policy Levy Costs Cost [£]' is a double-"Costs" stale column
    # that would appear alongside leaf PL columns in the export, misleading any manual
    # column sum. Drop it when leaf columns are available.
    existing_pl_leaf_cols = [col for col in pl_cost_cols if col in detailed_df.columns]
    if existing_pl_leaf_cols:
        detailed_df = detailed_df.drop(columns=['Policy Levy Costs Cost [£]'], errors='ignore')

    detailed_order = (
        [
            'DUoS Volumetric Cost [£]',
            'DUoS Fixed Cost [£]',
            'DUoS Capacity Cost [£]',
            'DUoS Reactive Cost [£]',
            'DUoS Exceeded Capacity Cost [£]',
            'Wholesale Commodity Cost [£]',
            'Policy Levy Costs [£]'
        ] +
        pl_cost_cols +
        [
            'Miscellaneous Charges Cost [£]'
        ] +
        misc_cost_cols +
        [
            'BSUoS Cost [£]',
            'Capacity Market Cost [£]',
            'Est. TNUoS Cost [£]',
            'Est. TNUoS Locational Cost [£]',
            'Est. TNUoS Residual Cost [£]',
            'Total Cost incl. VAT [£]'
        ]
    )  # Detailed annual price breakdown columns; aggregate Policy Levy and Misc listed before individual components so both Generic and demand-profile modes sort correctly

    detailed_df = reorder_columns_for_export(
        detailed_df, detailed_order
    )
    return detailed_df

def build_compressed_price_breakdown(df):
    """Builds a compressed annual bill breakdown with only the top-level user-facing cost categories."""
    compressed_df = pd.DataFrame(index=df.index)

    duos_components = [
        col for col in [
            'DUoS Volumetric Cost [£]',
            'DUoS Fixed Cost [£]',
            'DUoS Capacity Cost [£]',
            'DUoS Reactive Cost [£]',
            'DUoS Exceeded Capacity Cost [£]'
        ] if col in df.columns
    ]  # Explicit DUoS annual cost components combined into one DUoS total

    if duos_components:
        compressed_df['DUoS Total Cost [£]'] = df[duos_components].sum(axis=1)

    if 'Wholesale Commodity Cost [£]' in df.columns:
        compressed_df['Wholesale Commodity Cost [£]'] = df['Wholesale Commodity Cost [£]']

    existing_pl_cost_cols = [
        col for col in pl_cost_cols if col in df.columns
    ]  # Individual Policy Levy cost columns present (demand-profile mode)

    if existing_pl_cost_cols:
        compressed_df['Policy Levy Costs [£]'] = df[
            existing_pl_cost_cols
        ].sum(axis=1)
    elif 'Policy Levy Costs [£]' in df.columns:
        compressed_df['Policy Levy Costs [£]'] = (
            df['Policy Levy Costs [£]']
        )  # Aggregate Policy Levy column used directly in Generic flat-load mode

    existing_misc_cost_cols = [
        col for col in misc_cost_cols if col in df.columns
    ]  # Individual miscellaneous cost columns present (demand-profile mode)

    if existing_misc_cost_cols:
        compressed_df['Miscellaneous Charges Cost [£]'] = df[
            existing_misc_cost_cols
        ].sum(axis=1)
    elif 'Miscellaneous Charges Cost [£]' in df.columns:
        compressed_df['Miscellaneous Charges Cost [£]'] = (
            df['Miscellaneous Charges Cost [£]']
        )  # Aggregate miscellaneous column used directly in Generic flat-load mode

    if 'BSUoS Cost [£]' in df.columns:
        compressed_df['BSUoS Cost [£]'] = df['BSUoS Cost [£]']

    if 'Capacity Market Cost [£]' in df.columns:
        compressed_df['Capacity Market Cost [£]'] = df['Capacity Market Cost [£]']

    if 'Est. TNUoS Total Cost [£]' in df.columns:
        compressed_df['TNUoS Total Cost [£]'] = df['Est. TNUoS Total Cost [£]']
    elif 'Est. TNUoS Cost [£]' in df.columns:
        compressed_df['TNUoS Total Cost [£]'] = df['Est. TNUoS Cost [£]']

    return compressed_df

def save_price_breakdown_workbook(
    compressed_df, detailed_df, workbook_path
):
    """Saves compressed and detailed price breakdown tables to a two-sheet workbook."""
    with pd.ExcelWriter(workbook_path, engine='openpyxl') as writer:
        compressed_df.to_excel(
            writer, sheet_name='compressed_breakdown'
        )
        detailed_df.to_excel(
            writer, sheet_name='detailed_breakdown'
        )

# =============================================================================
# LOAD ACTIVE DEMAND PROFILE
# =============================================================================

if use_demand_profile:
    demand_profile_path = f"Input/Case Study Demand Profiles/{demand_profile_file}"
    if not os.path.exists(demand_profile_path):
        raise FileNotFoundError(
            f"Demand profile file not found: '{demand_profile_file}'. "
            f"Expected location: '{demand_profile_path}'."
        )
    active_profile_structured, active_profile_fin, active_usage_cols = (
        load_active_demand_profile(demand_profile_path)
    )
else:
    active_profile_structured = pd.DataFrame()
    active_profile_fin        = pd.DataFrame()
    active_usage_cols         = []

if use_demand_profile:
    demand_breakdown_export_order = (
        [f"{col} Cost [£]" for col in active_usage_cols] +
        ['Total Cost [£]']
    )  # Demand-type export order derived dynamically from the active demand profile component columns
else:
    demand_breakdown_export_order = []

price_component_cols = (
    ['DUoS Volumetric Cost [£]', 'Wholesale Commodity Cost [£]'] +
    [f'{col} Cost [£]' for col in pl_tariff_cols] +
    [f'{col} Cost [£]' for col in misc_charge_cols] +
    ['BSUoS Cost [£]', 'Capacity Market Cost [£]']
)  # Annual variable bill component labels used in detailed price breakdown outputs, including all Policy Levy and Miscellaneous charge components dynamically

pl_cost_cols = [f'{col} Cost [£]' for col in pl_tariff_cols]  # Detailed annual Policy Levy cost columns derived dynamically from the tariff workbook

misc_cost_cols = [f'{col} Cost [£]' for col in misc_charge_cols]  # Detailed annual miscellaneous charge cost columns derived dynamically from the miscellaneous workbook

PRICE_LABEL_MAP = {
    'ro Cost [£]':                       'Renewable Obligation [£]',
    'aahedc Cost [£]':                   'AAHEDC [£]',
    'fit Cost [£]':                      'Feed-in Tariff [£]',
    'ncc Cost [£]':                      'Network Charging Compensation [£]',
    'nrab Cost [£]':                     'Nuclear Regulated Asset Base [£]',
    'cfd Cost [£]':                      'Contract for Difference [£]',
    'Capacity Market Cost [£]':          'Capacity Market [£]',
    'managament fee Cost [£]':           'Management Fee [£]',
    'other Cost [£]':                    'Other Charges [£]',
    'shape Cost [£]':                    'Shape Charges [£]',
    'imbalance Cost [£]':                'Imbalance Charges [£]',
    'ccl Cost [£]':                      'Climate Change Levy [£]',
    'DUoS Volumetric Cost [£]':          'DUoS Volumetric [£]',
    'DUoS Fixed Cost [£]':               'DUoS Fixed [£]',
    'DUoS Capacity Cost [£]':            'DUoS Capacity [£]',
    'DUoS Reactive Cost [£]':            'DUoS Reactive [£]',
    'DUoS Exceeded Capacity Cost [£]':   'DUoS Exceeded Capacity [£]',
    'Wholesale Commodity Cost [£]':      'Wholesale Commodity [£]',
    'BSUoS Cost [£]':                    'BSUoS [£]',
    'Est. TNUoS Total Cost [£]':         'TNUoS Total [£]',
    'Est. TNUoS Cost [£]':               'TNUoS [£]',
    'Est. TNUoS Locational Cost [£]':    'TNUoS Locational [£]',
    'Est. TNUoS Residual Cost [£]':      'TNUoS Residual [£]',
}  # Long-form display labels for detailed price-breakdown columns, used by the disaggregated Price Breakdown chart

PRICE_BREAKDOWN_DISAGG_ORDER = [
    'Wholesale Commodity [£]',
    'DUoS Volumetric [£]', 'DUoS Fixed [£]', 'DUoS Capacity [£]',
    'DUoS Reactive [£]', 'DUoS Exceeded Capacity [£]',
    'BSUoS [£]',
    'TNUoS Total [£]', 'TNUoS [£]', 'TNUoS Locational [£]', 'TNUoS Residual [£]',
    'Renewable Obligation [£]', 'AAHEDC [£]', 'Feed-in Tariff [£]',
    'Network Charging Compensation [£]', 'Nuclear Regulated Asset Base [£]',
    'Contract for Difference [£]', 'Capacity Market [£]',
    'Management Fee [£]', 'Other Charges [£]', 'Shape Charges [£]',
    'Imbalance Charges [£]', 'Climate Change Levy [£]',
]  # Same top-level group order as price_breakdown_agg (Wholesale, Network, Policy Levy incl. Capacity Market, Misc), with each group's components kept sequential

misc_duos_cols = [
    'DUoS Fixed Cost [£]',
    'DUoS Capacity Cost [£]',
    'DUoS Reactive Cost [£]',
    'DUoS Exceeded Capacity Cost [£]'
]  # Explicit non-volumetric DUoS annual cost components preserved in outputs

price_breakdown_export_order = [
    'DUoS Volumetric Cost [£]',
    'DUoS Fixed Cost [£]',
    'DUoS Capacity Cost [£]',
    'DUoS Reactive Cost [£]',
    'DUoS Exceeded Capacity Cost [£]',
    'Wholesale Commodity Cost [£]',
    'Policy Levy Costs [£]',
    'Miscellaneous Charges Cost [£]',
    'BSUoS Cost [£]',
    'Capacity Market Cost [£]',
    'Est. TNUoS Cost [£]',
    'Est. TNUoS Locational Cost [£]',
    'Est. TNUoS Residual Cost [£]',
    'Est. TNUoS Total Cost [£]',
    'Total Cost [£]',
    'Total Cost incl. VAT [£]'
]  # Preferred export column order for detailed annual price breakdown outputs

demand_breakdown_export_order = []


# =============================================================================
# PLOTTING HELPER FUNCTIONS
# =============================================================================

def reformat(temp_day):
    """
    Reshapes a raw Average Template Days workbook into a tidy DataFrame.

    Input  : raw workbook with a merged header row (row 0) and data from row 1
    Output : DataFrame with Time and Year columns followed by monthly average
             price columns, ready for line-chart plotting
    """
    arrange  = temp_day.iloc[1:, 1:3].reset_index(drop=True)
    arrange1 = temp_day.iloc[1:, 3:].reset_index(drop=True)
    arrange1.columns = temp_day.iloc[0, 3:]
    return pd.concat([arrange, arrange1], axis=1)


def pie_chart(data, label, analysis_type, dno_reg, current_band, current_vol_type, analysis_root):
    """
    Saves one pie chart per year for a single DNO region showing bill category
    proportions; charts are written to analysis_root/Plots/Bill Breakdown by Year/.

    Parameters
    ----------
    data             : DataFrame   cost table with MultiIndex (DNO Region, Year) and
                                   a 'Total Cost [£]' column
    label            : str         chart title label (e.g. 'Price Type Breakdown')
    analysis_type    : str         analysis mode label used in the chart subtitle
    dno_reg          : str         short DNO region name used as subfolder and title
    current_band     : str         residual band label used in the chart subtitle
    current_vol_type : str         voltage/connection type label used in the chart subtitle
    analysis_root    : str         root directory for financial analysis outputs
    """
    pie_tab = data.xs(dno_reg, level='DNO Region', drop_level=False)

    def pct_label(pct):
        return f'{pct:.1f}%' if pct >= 2.0 else ''

    for yr in sorted(pie_tab.index.get_level_values('Year')):
        row = pie_tab.xs(
            yr, level='Year', drop_level=False
        ).reset_index(drop=True).loc[0].copy()

        total_cost = row['Total Cost [£]']
        row = row.drop([c for c in row.index if c.startswith('Total Cost')])

        if 'Capacity Market Cost [£]' in row.index:
            row['Policy Levy Costs [£]'] = (
                row.get('Policy Levy Costs [£]', 0)
                + row.get('Capacity Market Cost [£]', 0)
            )
            row = row.drop('Capacity Market Cost [£]')

        row = row.rename({
            'Policy Levy Costs [£]': 'Policy Levy Costs Total [£]',
            'Miscellaneous Charges Cost [£]': 'Miscellaneous Charges Total Cost [£]'
        })
        row = row[row > 0]
        _pie_order = [
            'Wholesale Commodity Cost [£]',
            'DUoS Total Cost [£]',
            'BSUoS Cost [£]',
            'TNUoS Total Cost [£]',
            'Policy Levy Costs Total [£]',
            'Miscellaneous Charges Total Cost [£]',
        ]
        _ordered = [c for c in _pie_order if c in row.index]
        _rest    = [c for c in row.index if c not in _pie_order]
        row = row[_ordered + _rest]

        fig, ax = plt.subplots(figsize=(8, 6))
        ax.pie(
            row,
            labels=row.index,
            autopct=pct_label,
            shadow=False,
            startangle=90,
            counterclock=False
        )
        fig.text(
            0.01, 0.99,
            f"Total Annual Cost: £{total_cost:,.0f}",
            fontsize=10, fontweight='bold', va='top', ha='left',
            bbox=dict(
                boxstyle='round,pad=0.4',
                facecolor='white', edgecolor='#333333', linewidth=1.2
            )
        )
        ax.set_title(
            f"{dno_reg}: {label} {yr}\n"
            f"{analysis_type} — {current_vol_type} | {current_band}"
        )

        save_dir = (
            f"{analysis_root}/Plots/"
            f"Bill Breakdown by Year/{dno_reg}"
        )
        os.makedirs(save_dir, exist_ok=True)
        fig.savefig(
            f"{save_dir}/{str(yr)[:4]}.png",
            dpi=300, bbox_inches='tight'
        )
        plt.close()


def bill_breakdown_stacked_line(data, label, analysis_type, dno_reg, current_band, current_vol_type, analysis_root):
    """
    Saves one stacked-area chart per DNO region showing bill category costs
    across all years together; charts are written to
    analysis_root/Plots/Bill Breakdown by Year/{dno_reg}/All Years Stacked.png.

    Parameters
    ----------
    data             : DataFrame   cost table with MultiIndex (DNO Region, Year)
    label            : str         chart title label (e.g. 'Price Type Breakdown')
    analysis_type    : str         analysis mode label used in the chart subtitle
    dno_reg          : str         short DNO region name used as subfolder and title
    current_band     : str         residual band label used in the chart subtitle
    current_vol_type : str         voltage/connection type label used in the chart subtitle
    analysis_root    : str         root directory for financial analysis outputs
    """
    region_tab = data.xs(dno_reg, level='DNO Region', drop_level=False)
    years = sorted(region_tab.index.get_level_values('Year'))
    rows = {}
    for yr in years:
        row = region_tab.xs(yr, level='Year', drop_level=False).reset_index(drop=True).loc[0].copy()
        row = row.drop([c for c in row.index if c.startswith('Total Cost')])
        if 'Capacity Market Cost [£]' in row.index:
            row['Policy Levy Costs [£]'] = (
                row.get('Policy Levy Costs [£]', 0)
                + row.get('Capacity Market Cost [£]', 0)
            )
            row = row.drop('Capacity Market Cost [£]')
        rows[str(yr)[:4]] = row
    plot_df = pd.DataFrame(rows).T
    plot_df = plot_df.loc[:, (plot_df > 0).any()]

    fig, ax = plt.subplots(figsize=(10, 6))
    plot_df.plot(kind='area', stacked=True, ax=ax)
    ax.yaxis.set_major_formatter(mpl.ticker.FuncFormatter(lambda x, p: f'{x:,.0f}'))
    ax.set_xlabel('Year')
    ax.set_ylabel('Annual Cost [£]')
    ax.set_title(
        f"{dno_reg}: {label} — All Years\n"
        f"{analysis_type} — {current_vol_type} | {current_band}"
    )
    ax.legend(loc='upper center', bbox_to_anchor=(0.5, -0.15),
              ncol=min(len(plot_df.columns), 4), fontsize=8)
    save_dir = f"{analysis_root}/Plots/Bill Breakdown by Year/{dno_reg}"
    os.makedirs(save_dir, exist_ok=True)
    fig.savefig(f"{save_dir}/All Years Stacked.png", dpi=150, bbox_inches='tight')
    plt.close()


def stacked_plot(data, label, analysis_type, current_band, current_vol_type, analysis_root, hex_dict_full):
    """
    Saves one stacked bar chart per year comparing annual costs across all DNO
    regions; charts are written to analysis_root/Plots/Regional Comparison/.

    Parameters
    ----------
    data             : DataFrame   cost table with MultiIndex (DNO Region, Year)
    label            : str         subfolder name and chart title element
    analysis_type    : str         analysis mode label; controls y-axis ceiling
    current_band     : str         residual band label for the chart subtitle
    current_vol_type : str         voltage/connection type label for the chart subtitle
    analysis_root    : str         root directory for financial analysis outputs
    hex_dict_full    : dict        region → hex colour mapping for 'Colour Visualisation' mode
    """
    plot_data = data.copy()
    plot_data.index = pd.MultiIndex.from_tuples(
        plot_data.index, names=['DNO Region', 'Year']
    )
    y_lim = data['Total Cost [£]'].max() * 1.20

    for yr in sorted(plot_data.index.get_level_values('Year').unique()):
        bar_tab    = plot_data.xs(
            yr, level='Year', drop_level=False
        ).sort_values('Total Cost [£]').reset_index(
            level='Year', drop=True
        )
        plot_df    = bar_tab.drop(columns=[c for c in bar_tab.columns if c.startswith('Total Cost')])

        plt.figure()
        if label == 'Colour Visualisation':
            ax         = plt.gca()
            reg_colors = [hex_dict_full.get(str(r), '#888888') for r in bar_tab.index]
            ax.bar(range(len(bar_tab)), bar_tab['Total Cost [£]'], color=reg_colors)
            ax.set_xticks(range(len(bar_tab)))
            ax.set_xticklabels(bar_tab.index, rotation=45, ha='right', fontsize=8)
            ax.set_ylim(0, y_lim)
            ax.yaxis.set_major_formatter(mpl.ticker.FuncFormatter(lambda x, p: f'{x:,.0f}'))
            all_reg_plot = ax
        else:
            all_reg_plot = plot_df.plot(
                kind='bar', stacked=True,
                legend=list(plot_df.columns),
                ylim=(0, y_lim)
            )
            all_reg_plot.set_xticklabels(
                all_reg_plot.get_xticklabels(),
                rotation=45, ha='right', fontsize=8
            )

        if label != 'Colour Visualisation':
            all_reg_plot.legend(loc='upper center', bbox_to_anchor=(0.5, -0.25),
                                ncol=min(len(plot_df.columns), 4), fontsize=8)
        all_reg_plot.set_ylabel("Annual Cost [£]")
        all_reg_plot.set_title(
            f"Annual Costs {label} - {yr}\n"
            f"{analysis_type} — {current_vol_type} | {current_band}"
        )

        save_dir = (
            f"{analysis_root}/Plots/"
            f"Regional Comparison/{label}"
        )
        os.makedirs(save_dir, exist_ok=True)
        all_reg_plot.get_figure().savefig(
            f"{save_dir}/{yr}.png",
            dpi=300, bbox_inches='tight'
        )
        plt.close('all')


def avs_stacked_plot(data, label, analysis_type, current_band, current_vol_type, analysis_root, hex_dict_full):
    """
    Saves a single average-across-all-years regional bar chart to
    analysis_root/Plots/Regional Comparison/ alongside the per-year charts.

    Parameters
    ----------
    data             : DataFrame   cost table with MultiIndex (DNO Region, Year)
    label            : str         subfolder name and chart title element
    analysis_type    : str         analysis mode label; controls y-axis ceiling
    current_band     : str         residual band label for the chart subtitle
    current_vol_type : str         voltage/connection type label for the chart subtitle
    analysis_root    : str         root directory for financial analysis outputs
    hex_dict_full    : dict        region → hex colour mapping for 'Colour Visualisation' mode
    """
    plot_data   = data.copy()
    plot_data.index = pd.MultiIndex.from_tuples(
        plot_data.index, names=['DNO Region', 'Year']
    )
    y_lim       = data['Total Cost [£]'].max() * 1.20
    all_reg_tab = plot_data.groupby(level='DNO Region').mean()
    plot_df     = all_reg_tab.sort_values(
        'Total Cost [£]'
    ).drop(columns=[c for c in all_reg_tab.columns if c.startswith('Total Cost')])

    plt.figure()
    if label == 'Colour Visualisation':
        ax         = plt.gca()
        sorted_tab = all_reg_tab.sort_values('Total Cost [£]')
        reg_colors = [hex_dict_full.get(str(r), '#888888') for r in sorted_tab.index]
        ax.bar(range(len(sorted_tab)), sorted_tab['Total Cost [£]'], color=reg_colors)
        ax.set_xticks(range(len(sorted_tab)))
        ax.set_xticklabels(sorted_tab.index, rotation=45, ha='right', fontsize=8)
        ax.set_ylim(0, y_lim)
        ax.yaxis.set_major_formatter(mpl.ticker.FuncFormatter(lambda x, p: f'{x:,.0f}'))
        all_reg_plot = ax
    else:
        all_reg_plot = plot_df.plot(
            kind='bar', stacked=True,
            legend=list(plot_df.columns),
            ylim=(0, y_lim)
        )
        all_reg_plot.set_xticklabels(
            all_reg_plot.get_xticklabels(),
            rotation=45, ha='right', fontsize=8
        )

    all_reg_plot.set_ylabel("Annual Cost [£]")
    all_reg_plot.set_title(
        f"Av. Annual Costs {label}\n"
        f"{analysis_type} — {current_vol_type} | {current_band}"
    )
    if label != 'Colour Visualisation':
        all_reg_plot.legend(loc='upper center', bbox_to_anchor=(0.5, -0.25),
                            ncol=min(len(plot_df.columns), 4), fontsize=8)

    save_dir = (
        f"{analysis_root}/Plots/"
        f"Regional Comparison/{label}"
    )
    os.makedirs(save_dir, exist_ok=True)
    all_reg_plot.get_figure().savefig(
        f"{save_dir}/Average All Years.png",
        dpi=300, bbox_inches='tight'
    )
    plt.close('all')


def colour_vis_trend_plot(data, analysis_type, current_band, current_vol_type, analysis_root, hex_dict_full):
    """
    Supplements the Colour Visualisation bar charts with a single line chart
    showing each DNO region's total annual cost trend across all years.
    Saved to National Trend/ so all trend charts are co-located.
    """
    plot_data = data.copy()
    plot_data.index = pd.MultiIndex.from_tuples(plot_data.index, names=['DNO Region', 'Year'])
    fig, ax = plt.subplots(figsize=(12, 6))
    for region in plot_data.index.get_level_values('DNO Region').unique():
        rd = plot_data.xs(region, level='DNO Region')
        years = sorted(rd.index)
        costs = [rd.loc[yr, 'Total Cost [£]'] for yr in years]
        color = hex_dict_full.get(str(region), '#888888')
        ax.plot([str(yr)[:4] for yr in years], costs,
                color=color, marker='o', label=region, linewidth=1.5)
    ax.yaxis.set_major_formatter(mpl.ticker.FuncFormatter(lambda x, p: f'{x:,.0f}'))
    ax.set_xlabel('Financial Year')
    ax.set_ylabel('Annual Cost [£]')
    ax.set_title(f"Annual Cost Trend by DNO Region\n{analysis_type} — {current_vol_type} | {current_band}")
    ax.legend(loc='upper center', bbox_to_anchor=(0.5, -0.15), ncol=7, fontsize=7)
    save_dir = f"{analysis_root}/Plots/National Trend"
    os.makedirs(save_dir, exist_ok=True)
    fig.savefig(f"{save_dir}/Total Cost Trend.png", dpi=150, bbox_inches='tight')
    plt.close()


def season_trend_plot(data, analysis_type, current_band, current_vol_type, analysis_root):
    """
    Stacked-area chart showing region-averaged winter/summer cost trends across all years.
    Saved to National Trend/ so all trend charts are co-located.
    """
    plot_data = data.copy()
    plot_data.index = pd.MultiIndex.from_tuples(plot_data.index, names=['DNO Region', 'Year'])
    cost_cols = [c for c in plot_data.columns if 'Total Cost' not in c]
    yearly_avg = plot_data[cost_cols].groupby(level='Year').mean()
    fig, ax = plt.subplots(figsize=(10, 6))
    yearly_avg.plot(kind='area', stacked=True, ax=ax)
    ax.yaxis.set_major_formatter(mpl.ticker.FuncFormatter(lambda x, p: f'{x:,.0f}'))
    ax.set_xlabel('Financial Year')
    ax.set_ylabel('Annual Cost [£]')
    ax.set_title(f"Seasonal Cost Trend — National Average\n{analysis_type} — {current_vol_type} | {current_band}")
    ax.legend(loc='upper center', bbox_to_anchor=(0.5, -0.15),
              ncol=min(len(cost_cols), 5), fontsize=8)
    save_dir = f"{analysis_root}/Plots/National Trend"
    os.makedirs(save_dir, exist_ok=True)
    fig.savefig(f"{save_dir}/Winter Vs Summer.png", dpi=150, bbox_inches='tight')
    plt.close()


def national_trend_plot(data, label, analysis_type, current_band, current_vol_type, analysis_root,
                        footnote_text=None):
    """
    Stacked area charts showing the national (region-averaged) cost trend over years.
    Produces two files per label: an absolute stacked area chart (£) and a 100% stacked
    area chart showing component proportions over time. 'Total Cost [£]' excluded from segments.
    Saved to National Trend/ alongside the other trend charts.
    """
    plot_data = data.copy()
    if not isinstance(plot_data.index, pd.MultiIndex):
        plot_data.index = pd.MultiIndex.from_tuples(
            plot_data.index, names=['DNO Region', 'Year']
        )
    elif None in plot_data.index.names:
        plot_data.index.names = ['DNO Region', 'Year']

    cost_cols = [c for c in plot_data.columns if 'Total Cost' not in c]
    yearly_avg = plot_data[cost_cols].groupby(level='Year').mean()
    years_sorted = sorted(yearly_avg.index, key=lambda y: int(str(y)[:4]))
    yearly_avg = yearly_avg.loc[years_sorted].fillna(0)
    x_pos    = list(range(len(years_sorted)))
    x_labels = [str(y)[:4] for y in years_sorted]

    n_cols = len(cost_cols)
    cmap   = mpl.colormaps.get_cmap('tab20' if n_cols > 10 else 'tab10')
    colors = [cmap(i / max(n_cols - 1, 1)) for i in range(n_cols)]
    save_dir   = f"{analysis_root}/Plots/National Trend"
    os.makedirs(save_dir, exist_ok=True)
    safe_label = label.replace('/', '-').replace(':', ',')

    def _fmt_axes(ax, ylabel, title):
        ax.set_xticks(x_pos)
        ax.set_xticklabels(x_labels, rotation=45, ha='right', fontsize=8)
        ax.set_xlabel('Financial Year')
        ax.set_ylabel(ylabel)
        ax.set_title(title)
        ax.legend(loc='upper center', bbox_to_anchor=(0.5, -0.20),
                  ncol=min(len(cost_cols), 5), fontsize=7)
        if footnote_text:
            ax.text(0, -0.38, footnote_text, transform=ax.transAxes,
                    fontsize=6.5, va='top', color='#898781')

    # Separate positive and negative contributions per column
    pos_values = [np.maximum(yearly_avg[col].values.astype(float), 0.0) for col in yearly_avg.columns]
    neg_cols   = [(i, col) for i, col in enumerate(yearly_avg.columns)
                  if (yearly_avg[col] < 0).any()]

    # ── Absolute stacked area (£) ──────────────────────────────────────────
    fig, ax = plt.subplots(figsize=(12, 6))
    ax.stackplot(x_pos, pos_values, labels=list(yearly_avg.columns), colors=colors, alpha=0.85)
    if neg_cols:
        neg_vals   = [np.minimum(yearly_avg[col].values.astype(float), 0.0) for _, col in neg_cols]
        neg_colors = [colors[i] for i, _ in neg_cols]
        neg_polys  = ax.stackplot(x_pos, neg_vals, colors=neg_colors, alpha=0.6)
        for poly in neg_polys:
            poly.set_hatch('///')
        ax.axhline(0, color='#555555', linewidth=0.7, linestyle='--', zorder=4)
    ax.yaxis.set_major_formatter(mpl.ticker.FuncFormatter(lambda v, _: f'{v:,.0f}'))
    _fmt_axes(ax, 'Annual Cost [£]',
              f"National Cost Trend — {label}\n{analysis_type} — {current_vol_type} | {current_band}")
    fig.savefig(f"{save_dir}/{safe_label}.png", dpi=150, bbox_inches='tight')
    plt.close()

    # ── 100% stacked area (%) — based on gross positive costs ─────────────
    # Negatives are clipped to zero: chart shows where gross costs go,
    # avoiding undefined percentage shares when a component is a net credit.
    pos_df     = yearly_avg.clip(lower=0)
    row_totals = pos_df.sum(axis=1).replace(0, np.nan)
    pct_values = [
        (pos_df[col] / row_totals * 100).fillna(0).values.astype(float)
        for col in pos_df.columns
    ]
    pct_title = f"National Cost Breakdown (%) — {label}\n{analysis_type} — {current_vol_type} | {current_band}"
    if neg_cols:
        neg_names  = ', '.join(col for _, col in neg_cols)
        pct_title += f"\n(negative components excluded from % — {neg_names})"
    fig, ax = plt.subplots(figsize=(12, 6))
    ax.stackplot(x_pos, pct_values, labels=list(pos_df.columns), colors=colors, alpha=0.85)
    ax.set_ylim(0, 100)
    ax.yaxis.set_major_formatter(mpl.ticker.FuncFormatter(lambda v, _: f'{v:.0f}%'))
    _fmt_axes(ax, 'Share of Gross Annual Cost [%]', pct_title)
    fig.savefig(f"{save_dir}/{safe_label} (%).png", dpi=150, bbox_inches='tight')
    plt.close()


def duos_fixed_charge_trend_plot(
    annex_data_cache, years_in, run_dno_regions, current_vol_type, vol_type_root
):
    """
    One clustered bar chart per DNO region showing DUoS fixed charge [p/MPAN/day]
    over time. Pre-TCR years (FY < 2022) show a single grey bar (no band
    differentiation existed). Post-TCR years (FY >= 2022) show four bars per year,
    one per residual band, clustered around the year tick. A dashed divider sits
    between FY 2021 and FY 2022 to mark the TCR boundary.
    """
    all_bands    = ['Band 1', 'Band 2', 'Band 3', 'Band 4']
    band_colors  = ['#1f77b4', '#ff7f0e', '#2ca02c', '#d62728']
    tcr_year     = 2022
    bw_pre       = 0.5    # width of the single pre-TCR bar
    bw_post      = 0.15   # width of each post-TCR band bar
    # Offsets so 4 bars are centred around the year tick with small gaps
    offsets      = [-0.225, -0.075, 0.075, 0.225]
    save_dir     = f"{vol_type_root}/DUoS Fixed Charges"
    os.makedirs(save_dir, exist_ok=True)

    for region in run_dno_regions:
        dno_region = region[:-5]

        pre_tcr  = {}                          # {yr_int: fixed_charge}
        post_tcr = {b: {} for b in all_bands}  # {band: {yr_int: fixed_charge}}

        for year in years_in:
            yr_int = int(str(year)[:4])

            if yr_int < tcr_year:
                try:
                    d1, read_year = annex_data_cache[(region, year)]
                    vol = get_vol_row_label(current_vol_type, read_year, 'Band 1')
                    d3 = (d1.loc[vol].to_frame().T
                          if isinstance(d1.loc[vol], pd.Series) else d1.loc[vol])
                    d3.columns = d3.columns.str.lower()
                    fixed = d3.loc[:, d3.columns.str.contains('fixed')].iat[0, 0]
                    pre_tcr[yr_int] = float(fixed)
                except Exception:
                    pass
            else:
                for band in all_bands:
                    try:
                        d1, read_year = annex_data_cache[(region, year)]
                        vol = get_vol_row_label(current_vol_type, read_year, band)
                        d3 = (d1.loc[vol].to_frame().T
                              if isinstance(d1.loc[vol], pd.Series) else d1.loc[vol])
                        d3.columns = d3.columns.str.lower()
                        fixed = d3.loc[:, d3.columns.str.contains('fixed')].iat[0, 0]
                        post_tcr[band][yr_int] = float(fixed)
                    except Exception:
                        pass

        if not pre_tcr and not any(post_tcr.values()):
            continue

        all_yr_ints = sorted(set(list(pre_tcr.keys()) +
                                 [y for b in post_tcr.values() for y in b.keys()]))

        fig, ax = plt.subplots(figsize=(12, 6))

        # Pre-TCR: one wide grey bar per year
        pre_added = False
        for yr, val in sorted(pre_tcr.items()):
            ax.bar(yr, val, width=bw_pre, color='#888888',
                   label='All Bands (pre-TCR)' if not pre_added else '',
                   zorder=3)
            if val == 0:
                ax.plot(yr, 0, 'o', color='#888888', ms=5, zorder=4)
            pre_added = True

        # Post-TCR: four narrow bars clustered around each year tick
        band_added = set()
        for band, color, offset in zip(all_bands, band_colors, offsets):
            for yr, val in sorted(post_tcr[band].items()):
                ax.bar(yr + offset, val, width=bw_post, color=color,
                       label=band if band not in band_added else '',
                       zorder=3)
                if val == 0:
                    ax.plot(yr + offset, 0, 'o', color=color, ms=5, zorder=4)
                band_added.add(band)

        ax.set_xlim(min(all_yr_ints) - 0.6, max(all_yr_ints) + 0.6)
        ax.set_xticks(all_yr_ints)
        ax.set_xticklabels([str(y) for y in all_yr_ints])
        ax.yaxis.set_major_formatter(
            mpl.ticker.FuncFormatter(lambda x, p: f'{x:,.0f}')
        )
        ax.set_xlabel('Financial Year')
        ax.set_ylabel('Fixed Charge [p/MPAN/day]')
        ax.set_title(f'Fixed Charge: {dno_region}\n{current_vol_type}')
        ax.legend()

        fig.savefig(f"{save_dir}/{dno_region}.png", dpi=150, bbox_inches='tight')
        plt.close(fig)


def plot_ecb_trend(ecb_collection, all_vol_types, cross_vol_dir, partial_yr_set=None):
    """
    Saves a line chart to cross_vol_dir showing the region- and band-averaged
    Effective Cost Benchmark trend across years, one line per voltage type.
    Partial years are excluded from the chart; the ECB Excel retains all data.
    """
    partial_yr_set = partial_yr_set or set()
    last_updated   = datetime.date.today().strftime('%d %b %Y')

    # Determine last full year across the collection for the footnote
    all_yr_strs = sorted({str(y) for df in ecb_collection.values() for y in df.columns},
                         key=lambda y: int(y))
    full_yr_strs = [y for y in all_yr_strs if y not in partial_yr_set]
    last_full_yr = f"FY{full_yr_strs[-1]}" if full_yr_strs else None

    fig, ax = plt.subplots(figsize=(10, 6))
    for vt in all_vol_types:
        band_series = [
            df.drop(columns=[c for c in df.columns if str(c) in partial_yr_set],
                    errors='ignore').mean(axis=0)
            for (v, b), df in ecb_collection.items() if v == vt
        ]
        if not band_series:
            continue
        vt_trend = pd.concat(band_series, axis=1).mean(axis=1)
        years_sorted = sorted(vt_trend.index, key=lambda y: int(str(y)))
        ax.plot([str(y) for y in years_sorted], vt_trend.loc[years_sorted].values,
                marker='o', label=vt)
    ax.set_xlabel('Year')
    ax.set_ylabel('Effective Cost [p/kWh]')
    ax.set_title('Effective Cost Benchmark Trend by Voltage Connection')
    ax.grid(True, linestyle='--', linewidth=0.5, alpha=0.5)
    ax.set_axisbelow(True)
    ax.legend(loc='upper center', bbox_to_anchor=(0.5, -0.10),
              ncol=len(all_vol_types), fontsize=8)
    footnote = f"Last updated: {last_updated}"
    if last_full_yr:
        footnote += f"  |  Last full year: {last_full_yr}"
    ax.text(0, -0.18, footnote, transform=ax.transAxes,
            fontsize=6.5, va='top', color='#898781')
    os.makedirs(cross_vol_dir, exist_ok=True)
    fig.savefig(f"{cross_vol_dir}/Effective Cost Benchmark Trend.png", dpi=150, bbox_inches='tight')
    plt.close()


# =============================================================================
# MAIN EXECUTION LOOP
# =============================================================================

vol_llf_map = {
    'LV':                      'low voltage network',
    'LV Sub':                  'low voltage substation',
    'HV':                      'high voltage network',
    'Non-Domestic Aggregated': 'low voltage network'
}

def quarterly_cost_plot(quarterly_df, analysis_type, current_band, current_vol_type, analysis_root):
    """
    Stacked bar chart of costs broken down by FY quarter (Q1=Apr–Jun,
    Q2=Jul–Sep, Q3=Oct–Dec, Q4=Jan–Mar) and compressed cost category.
    Volumetric costs are quarterly sums from HH data; fixed annual charges
    (DUoS and TNUoS) are pre-divided by 4 before reaching this function.
    Produces one chart per DNO region and a region-averaged national chart
    saved alongside the other national trend plots.
    Partial quarters are excluded from charts; Excel retains all data.
    """
    cost_cols = ['Wholesale [£]', 'DUoS [£]', 'BSUoS [£]', 'TNUoS [£]',
                 'Policy Levy Costs [£]', 'Miscellaneous [£]', 'VAT [£]']
    cost_cols = [c for c in cost_cols if c in quarterly_df.columns]

    # Validated categorical palette: blue→aqua→yellow→green→violet→red→gray(VAT)
    _palette   = ['#2a78d6', '#1baf7a', '#eda100', '#008300', '#4a3aa7', '#e34948', '#999999']
    col_color  = {c: _palette[i % len(_palette)] for i, c in enumerate(cost_cols)}

    # Compute footnote strings from the full (unfiltered) df once
    _last_updated = datetime.date.today().strftime('%d %b %Y')
    _full_q = quarterly_df[~quarterly_df.get('Partial', pd.Series(False, index=quarterly_df.index)).fillna(False)]
    if not _full_q.empty:
        _lq = _full_q.sort_values(['Year', 'Quarter']).iloc[-1]
        _last_full_q = f"{_lq['Quarter']} FY{_lq['Year']}"
    else:
        _last_full_q = None
    _footnote = f"Last updated: {_last_updated}"
    if _last_full_q:
        _footnote += f"  |  Last full quarter: {_last_full_q}"

    def _draw(df, title, save_path):
        df = df.copy().sort_values(['Year', 'Quarter'])
        # Exclude partial quarters from the chart
        if 'Partial' in df.columns:
            df = df[~df['Partial'].fillna(False).astype(bool)]
        if df.empty:
            return
        df['Label'] = df['Year'].astype(str) + ' ' + df['Quarter']
        n   = len(df)
        fig, ax = plt.subplots(figsize=(max(14, n * 0.55), 6))

        bottom = np.zeros(n)
        x      = np.arange(n)
        for col in cost_cols:
            vals = df[col].fillna(0).values
            ax.bar(x, vals, bottom=bottom, width=0.75,
                   color=col_color[col], label=col, zorder=3)
            bottom += vals

        # Hairline dividers between year groups
        years = df['Year'].astype(str).values
        for i in range(1, n):
            if years[i] != years[i - 1]:
                ax.axvline(x=i - 0.5, color='#e1e0d9', linewidth=0.8, zorder=2)

        ax.set_xticks(x)
        ax.set_xticklabels(df['Label'].values, rotation=90, fontsize=7)
        ax.yaxis.set_major_formatter(
            mpl.ticker.FuncFormatter(lambda v, _: f'{v:,.0f}')
        )
        ax.set_xlabel('Financial Year / Quarter', labelpad=8)
        ax.set_ylabel('Cost [£]')
        ax.set_title(title)
        ax.legend(loc='upper center', bbox_to_anchor=(0.5, -0.22),
                  ncol=len(cost_cols), fontsize=8)
        ax.text(0, -0.38, _footnote, transform=ax.transAxes,
                fontsize=6.5, va='top', color='#898781')
        os.makedirs(os.path.dirname(save_path), exist_ok=True)
        fig.savefig(save_path, dpi=150, bbox_inches='tight')
        plt.close(fig)

    # Per-region charts
    qb_dir = f"{analysis_root}/Plots/Quarterly Breakdown"
    for region in sorted(quarterly_df['DNO Region'].unique()):
        _draw(
            quarterly_df[quarterly_df['DNO Region'] == region],
            f"Quarterly Cost Breakdown: {region}\n"
            f"{analysis_type} — {current_vol_type} | {current_band}",
            f"{qb_dir}/{region}.png"
        )

    # National average: mean across regions per (year, quarter)
    nat_df = quarterly_df.groupby(['Year', 'Quarter'])[cost_cols].mean().reset_index()
    if 'Partial' in quarterly_df.columns:
        nat_df = nat_df.merge(
            quarterly_df.groupby(['Year', 'Quarter'])['Partial'].any().reset_index(),
            on=['Year', 'Quarter']
        )
    _draw(
        nat_df,
        f"Quarterly Cost Breakdown — National Average\n"
        f"{analysis_type} — {current_vol_type} | {current_band}",
        f"{analysis_root}/Plots/National Trend/Quarterly Cost Breakdown.png"
    )


def _place_right_labels(ax, endpoints, x_text=48.5, min_gap_frac=0.04):
    """
    Place right-margin line annotations without overlap.

    endpoints : list of (y_val, label_str, color)
    Labels are sorted by y_val, spread vertically with minimum gap =
    min_gap_frac * y_range, and connected to their true line endpoints
    with a thin leader line.
    """
    y_min, y_max = ax.get_ylim()
    min_gap = (y_max - y_min) * min_gap_frac

    sorted_ep = sorted(endpoints, key=lambda t: t[0])
    adjusted  = [ep[0] for ep in sorted_ep]

    # Forward pass: push upward when labels are too close
    for i in range(1, len(adjusted)):
        if adjusted[i] - adjusted[i - 1] < min_gap:
            adjusted[i] = adjusted[i - 1] + min_gap

    # Backward pass: pull downward if forward pass overflowed the top
    for i in range(len(adjusted) - 2, -1, -1):
        if adjusted[i + 1] - adjusted[i] < min_gap:
            adjusted[i] = adjusted[i + 1] - min_gap

    y_range = y_max - y_min
    for (y_true, label, color), y_adj in zip(sorted_ep, adjusted):
        displaced = abs(y_adj - y_true) > y_range * 0.005
        if displaced:
            ax.annotate(
                label,
                xy=(47, y_true),
                xytext=(x_text, y_adj),
                va='center', fontsize=8, color=color,
                annotation_clip=False,
                arrowprops=dict(
                    arrowstyle='-', color=color, lw=0.4,
                    connectionstyle='arc3,rad=0'
                )
            )
        else:
            ax.text(
                x_text, y_adj, label,
                va='center', fontsize=8, color=color,
                transform=ax.transData, clip_on=False
            )


ecb_collection = {}  # (vol_type, band) → annual_p_per_kwh_df; populated in Generic mode only

for current_vol_type in run_vol_types:

    print(f"\n{'='*60}")
    print(f"Processing Voltage Type: {current_vol_type}")
    print(f"{'='*60}")

    vol_llf = vol_llf_map.get(current_vol_type)

    # LLFs vary by voltage type and region only — not by band.
    # Compute once per voltage type and save reference files at the voltage-type level.
    if run_mode == 'Generic':
        vol_type_root = f"{output_root}/{current_vol_type}"
    else:
        vol_type_root = output_root
    llfprices_full, dftime_tup_out = llf(
        current_vol_type, vol_llf, run_dno_regions,
        years_past, vol_type_root,
        save_reference=(run_mode == 'Generic')
    )

    # Cache raw Annex 1 DUoS tariff DataFrames keyed by (region, year).
    # The same file is read for every band within this voltage type; only the
    # row extracted from it (via get_vol_row_label) differs by band.
    # Reading once here cuts Annex 1 I/O by 75% across a four-band run.
    annex_data_cache = {}
    for _region in run_dno_regions:
        for _year in years_in:
            _annex_sheet = get_annex_sheet(_year)
            try:
                _data = pd.read_excel(
                    f"Input/Network Prices/Distribution/{_year}/{_region}",
                    sheet_name=_annex_sheet
                )
                _read_year = _year
            except Exception:
                _annex_sheet = get_annex_sheet(latest_run_year)
                _data = pd.read_excel(
                    f"Input/Network Prices/Distribution/"
                    f"{latest_run_year}/{_region}",
                    sheet_name=_annex_sheet
                )
                _read_year = latest_run_year
            _d1 = _data.set_index(_data.columns[0])
            _d1.columns = _d1.loc['Tariff name']
            annex_data_cache[(_region, _year)] = (_d1, _read_year)

    if run_mode == 'Generic':
        flat_loads_df = pd.DataFrame([
            {
                'Band':           band,
                'Basis [kW]':     generic_flat_loads[current_vol_type][band],
                'Capacity [kVA]': round(
                    generic_flat_loads[current_vol_type][band] / power_factor, 4
                )
            }
            for band in generic_flat_loads.get(current_vol_type, {})
        ])
        n_reg = len(run_dno_regions)
        n_yr  = len(years_in)
        n_max = max(n_reg, n_yr)
        scope_df = pd.DataFrame({
            'DNO Region': [r[:-5] for r in run_dno_regions] + [''] * (n_max - n_reg),
            'Year':       [str(y) for y in years_in]         + [''] * (n_max - n_yr),
        })
        summary_path = f"{vol_type_root}/Reference Data/Generic Run Summary.xlsx"
        with pd.ExcelWriter(summary_path, engine='openpyxl') as writer:
            flat_loads_df.to_excel(writer, sheet_name='flat_loads', index=False)
            scope_df.to_excel(writer, sheet_name='scope', index=False)

    if not apply_vat:
        duos_fixed_charge_trend_plot(
            annex_data_cache, years_in, run_dno_regions,
            current_vol_type, vol_type_root
        )

    for current_band in bands_to_run:

        # ── Step 1: Basis & profile selection ─────────────────────────────────
        band_number = int(current_band[-1])
        print(f"\n  Band: {current_band}")

        if run_mode == 'Generic':
            current_basis  = generic_flat_loads[current_vol_type][current_band]
            capacity       = current_basis / power_factor
            reactive_power = current_basis * np.tan(np.arccos(power_factor))
        else:
            current_basis = basis

        if apply_vat:
            _daily_kwh = current_basis * 24
            vat_rate   = 0.05 if _daily_kwh <= 33.0 else 0.20
            _rate_str  = (
                f"5% (de minimis — {_daily_kwh:.1f} kWh/day ≤ 33)"
                if vat_rate == 0.05 else
                f"20% (standard — {_daily_kwh:.1f} kWh/day > 33)"
            )
            print(f"  VAT rate: {_rate_str}")

        # Price profiles always read from Generic Results regardless of run_mode;
        # Specific mode reads from Generic Results and writes financial outputs to Case Studies.
        price_profile_root = (
            f"{generic_root}/{current_vol_type}/{current_band}"
        )

        if run_mode == 'Generic':
            vol_root = price_profile_root  # Generic outputs stay organised by voltage type and residual band
        else:
            vol_root = output_root  # Specific outputs are written directly under the case-study root


        duos_methods = ['Actual']  # DUoS method is always 'Actual' — historical tariff data only
        print(f"  ✓ DUoS method: Actual (historical data only)")

        # ── Detect partial years from price profile parquet ───────────────────
        # Count non-NaN HH slots per year; a year with fewer than 17520 rows is partial.
        _probe_path = (
            f"{price_profile_root}/1 - Price Profiles/Half-Hourly Prices/"
            f"{run_dno_regions[0][:-5]}.parquet"
        )
        _probe_cols = [f'Year {y}' for y in years_in]
        _probe = pd.read_parquet(_probe_path, columns=_probe_cols)
        partial_year_days = {}
        for _y in years_in:
            _col = f'Year {_y}'
            if _col in _probe.columns:
                _n = int(_probe[_col].notna().sum())
                partial_year_days[str(_y)] = None if _n >= 17520 else max(1, _n // 48)
            else:
                partial_year_days[str(_y)] = None
        del _probe
        _partial_yrs    = {k: v for k, v in partial_year_days.items() if v is not None}
        _partial_yr_set = set(_partial_yrs.keys())  # string year keys, e.g. {'2026'}
        if _partial_yrs:
            print(f"  Partial years detected: {_partial_yrs}")

        # Footnote text used on chronological trend charts
        _last_updated = datetime.date.today().strftime('%d %b %Y')
        _full_years   = [y for y in sorted(years_in, key=int) if str(y) not in _partial_yr_set]
        _last_full_yr = str(_full_years[-1]) if _full_years else str(max(years_in))
        _footnote_annual = (
            f"Last updated: {_last_updated}  |  Last full year: FY{_last_full_yr}"
        )

        # ── Step 2: Additional DUoS charges (fixed / capacity / reactive) ──────
        # File reads are shared across bands via annex_data_cache; only the row
        # label returned by get_vol_row_label() differs between bands.
        add_charges_list = []
        for region in run_dno_regions:
            dno_region = region[:-5]
            for year in years_in:
                d1, read_year = annex_data_cache[(region, year)]
                vol = get_vol_row_label(current_vol_type, read_year, current_band)
                d3 = (
                    d1.loc[vol].to_frame().T
                    if isinstance(d1.loc[vol], pd.Series)
                    else d1.loc[vol]
                )
                d3.columns = d3.columns.str.lower()
                fixed      = d3.loc[:, d3.columns.str.contains(
                    "fixed")].iat[0, 0]
                cap        = d3.loc[:, d3.columns.str.contains(
                    "capacity"
                ) & ~d3.columns.str.contains(
                    "exceeded|excess")].iat[0, 0]
                exceed_cap = d3.loc[:, d3.columns.str.contains(
                    'exceeded capacity|excess capacity')].iat[0, 0]
                reactive   = d3.loc[:, d3.columns.str.contains(
                    "reactive")].iat[0, 0]
                add_charges_list.append({
                    'DNO Region':                           dno_region,
                    'Year':                                 year,
                    'Fixed charge [p/MPAN/day]':            fixed,
                    'Capacity charge [p/kVA/day]':          cap,
                    'Exceeded capacity charge [p/kVA/day]': exceed_cap,
                    'Reactive charge [p/kVArh]':            reactive,
                    'Days Covered':                         partial_year_days.get(str(year)) or 365,
                })

        add_charges = pd.DataFrame(add_charges_list).set_index('DNO Region')

        add_charges['DUoS Fixed Cost [£]'] = (
            add_charges['Fixed charge [p/MPAN/day]'] * (add_charges['Days Covered'] / 100)
        )  # DUoS fixed charge: p/MPAN/day × days covered → £ (pro-rated for partial years)

        add_charges['DUoS Capacity Cost [£]'] = (
            add_charges['Capacity charge [p/kVA/day]'] * capacity * (add_charges['Days Covered'] / 100)
        )  # DUoS capacity charge: p/kVA/day × kVA × days covered → £ (pro-rated for partial years)

        # DUoS reactive power charge excluded: power factor assumed ≥ 0.95 throughout,
        # so the excess-reactive threshold is never breached and the charge is zero.
        # See DUoS_Reactive_Power_Briefing.md for the methodological decision.
        add_charges['DUoS Reactive Cost [£]'] = 0.0

        add_charges1    = add_charges.set_index('Year', append=True)
        add_charges_out = add_charges1[[
            'DUoS Fixed Cost [£]',
            'DUoS Capacity Cost [£]',
            'DUoS Reactive Cost [£]'
        ]]  # Annual non-volumetric DUoS charge outputs merged into final financial tables

        # ── Step 3: Region & year cost calculation ────────────────────────────
        print(f"\n  --- Analysis: {analysis_label} ---")

        fin_root = f"{vol_root}/{financial_dir_name}"  # Financial-analysis root for the active run mode
        analysis_root = f"{fin_root}/{analysis_label}"  # Active analysis branch root inside the financial-analysis folder

        os.makedirs(analysis_root, exist_ok=True)  # Ensure the active analysis branch directory exists before saving run settings or other analysis outputs

        if run_mode == 'Specific':
            run_settings_tables = create_run_settings_tables(
                run_mode=run_mode,
                case_study_name=case_study_name,
                current_vol_type=current_vol_type,
                current_band=current_band,
                analysis_basis=analysis_basis,
                analysis_label=analysis_label,
                demand_profile_file=demand_profile_file,
                current_basis=current_basis,
                capacity=capacity,
                reactive_power=reactive_power,
                run_dno_regions=run_dno_regions,
                years_in=years_in,
                use_flat_load=use_flat_load,
                use_demand_profile=use_demand_profile
            )
            run_settings_path = f"{analysis_root}/run_settings.xlsx"
            with pd.ExcelWriter(
                run_settings_path, engine='openpyxl'
            ) as writer:
                for sheet_name, settings_df in run_settings_tables.items():
                    settings_df.to_excel(
                        writer, sheet_name=sheet_name, index=False
                    )

        # Chart root follows the active run-mode folder naming convention.
        chart_root = f"{vol_root}/{chart_dir_name}"  # Chart output root for the active run mode

        peak_base_cols = [
            'Year', f"{peak_label} Peak Vs NPeak", 'Non-Peak'
        ]
        win_sum_cols   = ['Year', 'Winter', 'Summer']

        active_costs_list        = []
        usecosts_list            = []
        peak_split_list          = []
        season_split_list        = []
        annual_profiles_cache    = {}
        quarterly_costs_list     = []
        _cs_band_breakdown_list  = []  # per (region, year) DUoS band breakdown frames
        _cs_key_metrics_rows     = []  # per (region, year) case-study scalar metrics
        _specific_annual_kwh     = None  # annual kWh from demand profile; set on first Specific mode year

        for region in run_dno_regions:
            dno_region = region[:-5]
            dno_no     = int(region[:2])

            # Stage 6: always read from price_profile_root (Generic Results)
            annual_profile_path = (
                f"{price_profile_root}/1 - Price Profiles/"
                f"Half-Hourly Prices/{region[:-5]}.parquet"
            )  # Full path to the HH price profile for the current region

            annual_profiles = pd.read_parquet(annual_profile_path)

            for yr in years_in:
                # Stage 5: always DUoS Actual — no LinReg/Inflation branch
                annual_profiles[f'DUoS [p/kWh] {yr}'] = (
                    annual_profiles[f'DUoS Actual [p/kWh] {yr}']
                )
                annual_profiles[f'Total [p/kWh] {yr}'] = (
                    annual_profiles[f'DUoS [p/kWh] {yr}'] +
                    annual_profiles[
                        f'Wholesale Commodity [p/kWh] {yr}'
                    ] +
                    annual_profiles[
                        f'Policy Levy Costs [p/kWh] {yr}'
                    ] +
                    annual_profiles[
                        f'Miscellaneous Charges [p/kWh] {yr}'
                    ] +
                    annual_profiles[
                        f'BSUoS [p/kWh] {yr}'
                    ] +
                    annual_profiles[
                        f'Capacity Market [p/kWh] {yr}'
                    ]
                )

            annual_profiles_cache[region] = annual_profiles.copy()
            reg_active_fin_list = []

            for yr in list(years_in):
                yr_str = str(yr)
                _hh_band_srs = None  # aligned band-label Series; set below when use_demand_profile
                _hh_rate_srs = None  # aligned DUoS rate Series; set below when use_demand_profile

                try:
                    llf_price_reg = llfprices_full.loc[
                        (yr_str, dno_region)
                    ].to_frame().dropna()
                except KeyError:
                    # Use the closest available year when this year has no LLF data
                    available_years = [
                        y for y, r in llfprices_full.index
                        if r == dno_region
                    ]
                    if available_years:
                        fallback_year = sorted(
                            available_years,
                            key=lambda x: abs(int(x) - int(yr_str))
                        )[0]
                        llf_price_reg = llfprices_full.loc[
                            (fallback_year, dno_region)
                        ].to_frame().dropna()
                    else:
                        llf_price_reg = pd.DataFrame()

                columns = (
                    [
                        f'Year {yr_str}',
                        f'DUoS [p/kWh] {yr_str}',
                        f'Wholesale Commodity [p/kWh] {yr_str}',
                        f'Policy Levy Costs [p/kWh] {yr_str}',
                    ] +
                    [f'{col} [p/kWh] {yr_str}' for col in pl_tariff_cols] +
                    [
                        f'Miscellaneous Charges [p/kWh] {yr_str}',
                        f'BSUoS [p/kWh] {yr_str}',
                        f'Capacity Market [p/kWh] {yr_str}',
                        f'Total [p/kWh] {yr_str}'
                    ]
                )
                # Include DUoS Band column if present (added to parquet in Step 1)
                _band_col_yr = f'DUoS Band {yr_str}'
                if _band_col_yr in annual_profiles.columns:
                    columns = columns + [_band_col_yr]
                year_prices = annual_profiles[columns].dropna(
                    how='all'
                ).copy()

                const_template_out = None
                const_out          = None

                if use_flat_load:
                    const_template  = year_prices.set_index(
                        f'Year {yr_str}'
                    )
                    # DUoS Band column is a string label — drop before price→cost scaling
                    if _band_col_yr in const_template.columns:
                        const_template = const_template.drop(columns=[_band_col_yr])
                    const_template1 = const_template * (
                        current_basis / (2 * 100)
                    )
                    const_template1.columns = (
                        const_template1.columns.str[:-13] + " Cost [£]"
                    )

                    # Explicitly rename DUoS to DUoS Volumetric
                    const_template1 = const_template1.rename(
                        columns={'DUoS Cost [£]': 'DUoS Volumetric Cost [£]'}
                    )

                    # Collect quarterly volumetric sums for quarterly cost plot.
                    # Resample to FY quarters (QS-APR: Q1=Apr, Q2=Jul, Q3=Oct, Q4=Jan).
                    _q_vol_cols = [
                        'DUoS Volumetric Cost [£]',
                        'Wholesale Commodity Cost [£]',
                        'Policy Levy Costs Cost [£]',  # renamed from [p/kWh] via str[:-13] + " Cost [£]"
                        'Miscellaneous Charges Cost [£]',
                        'BSUoS Cost [£]',
                        'Capacity Market Cost [£]',
                    ]
                    _q_src = const_template1[
                        [c for c in _q_vol_cols if c in const_template1.columns]
                    ]
                    _q_vol = _q_src.resample('QS-APR').sum()
                    # Count actual HH slots per quarter to detect partial quarters
                    _q_vol['_hh_count'] = _q_src.resample('QS-APR').count().iloc[:, 0]
                    _q_vol.index.name = 'Quarter Start'
                    _q_vol['DNO Region'] = dno_region
                    _q_vol['Year']       = yr
                    quarterly_costs_list.append(_q_vol)

                    const_template_out = const_template1.rename(
                        columns={'Total Cost [£]': 'Overall'}
                    )
                    const_ann = const_template1.sum().to_frame()
                    const_ann.columns = [dno_region]
                    const_out = const_ann.T.iloc[:, :-1]

                price_types = columns[1:-1]

                year_prices[f'Year {yr_str}'] = pd.to_datetime(
                    year_prices[f'Year {yr_str}']
                )

                template_cost = pd.DataFrame()

                if use_demand_profile:
                    template_date = template_dates_full[[
                        f'Year {yr_str}',
                        f'Date Indexes {yr_str}',
                        f'Month {yr_str}',
                        f'Day in Week {yr_str}',
                        f'Week in Month {yr_str}',
                        f'Time {yr_str}'
                    ]]
                    year_prices1 = template_date.merge(
                        year_prices, how='outer',
                        left_on=f'Year {yr_str}',
                        right_on=f'Year {yr_str}'
                    )
                    template_cost = year_prices1.merge(
                        active_profile_structured,
                        left_on=[f"Date Indexes {yr_str}"],
                        right_on=["Date Indexes"]
                    )

                store_temp_out = None

                if use_demand_profile:
                    columns_usage = [
                        f'Year {yr_str}',
                        f'Total [p/kWh] {yr_str}',
                        'Overall'
                    ] + active_usage_cols

                    template_usage = template_cost[columns_usage].copy()
                    template_usage[active_usage_cols] = (
                        template_usage[active_usage_cols].multiply(
                            template_usage[
                                f'Total [p/kWh] {yr_str}'
                            ] / (100 * 2),
                            axis='index'
                        )
                    )
                    store_temp_out = template_usage.set_index(
                        f'Year {yr_str}'
                    )[['Overall'] + active_usage_cols].astype(float)

                    if _specific_annual_kwh is None:
                        _specific_annual_kwh = float(
                            store_temp_out['Overall'].sum() * 0.5
                        )  # kWh for one profile year; used by effective-rate benchmark in Specific mode

                    # ── Case Study Analytics: extract band / rate from template_cost ──
                    # template_cost has the same row count and order as store_temp_out
                    # (store_temp_out is derived from template_cost). Using template_cost
                    # avoids the annual_profiles shape mismatch (annual_profiles contains
                    # all years wide; template_cost is already filtered to this year).
                    _band_col_tc = f'DUoS Band {yr_str}'
                    if _band_col_tc in template_cost.columns:
                        _hh_band_srs = template_cost[_band_col_tc].values
                    _hh_rate_srs = template_cost[f'DUoS [p/kWh] {yr_str}'].values

                    peak_split_list.append(
                        peak_base_calc(
                            store_temp_out, dno_region,
                            yr, peak_base_cols
                        )
                    )
                    season_split_list.append(
                        sum_winter_calc(
                            store_temp_out, dno_region,
                            yr, win_sum_cols
                        )
                    )

                    if active_usage_cols:
                        template_usage1 = (
                            template_usage[active_usage_cols].copy()
                        )
                        template_usage1.columns = [
                            f"{lbl} Cost [£]" for lbl in active_usage_cols
                        ]
                        usage_costs_out = pd.DataFrame(
                            template_usage1.sum(), columns=[dno_region]
                        ).T
                        usage_costs_out['Total Cost [£]'] = (
                            template_usage1.sum().sum()
                        )
                        usage_costs_out['Year'] = yr
                        usecosts_list.append(
                            usage_costs_out.set_index('Year', append=True)
                        )

                if use_flat_load and const_template_out is not None:
                    peak_split_list.append(
                        peak_base_calc(
                            const_template_out, dno_region,
                            yr, peak_base_cols
                        )
                    )
                    season_split_list.append(
                        sum_winter_calc(
                            const_template_out, dno_region,
                            yr, win_sum_cols
                        )
                    )

                ptype_costs_out = pd.DataFrame()

                if use_demand_profile:
                    columns_ptype   = (
                        [f'Year {yr_str}', 'Overall'] +
                        list(price_types)
                    )
                    template_p_type = template_cost[
                        columns_ptype
                    ].copy()
                    template_p_type.columns = (
                        [f'Year {yr_str}', 'Overall'] +
                        [
                            c.replace(f' [p/kWh] {yr_str}', ' Cost [£]')
                             .replace('DUoS Cost [£]', 'DUoS Volumetric Cost [£]')
                            for c in price_types
                        ]
                    )
                    template_p_type1 = (
                        template_p_type.iloc[:, 2:].multiply(
                            template_p_type['Overall'] / (100 * 2),
                            axis="index"
                        )
                    )
                    ptype_costs_out = pd.DataFrame(
                        template_p_type1.sum(), columns=[dno_region]
                    ).T
                    ptype_costs_out['Year'] = yr
                    ptype_costs_out = ptype_costs_out.set_index(
                        'Year', append=True
                    )

                    # ── Flat-levy cost correction ─────────────────────────────────────────
                    # The Date Index tuple merge can drop demand rows that have no matching
                    # (Month, DayInWeek, WeekInMonth, Time) in the price profile (e.g. a
                    # 5th weekday occurrence in a month), causing a spurious cost offset
                    # between baseline and TOU-shifted profiles despite identical total kWh.
                    # For flat or seasonal per-kWh levies (RO, FiT, CfD, AAHEDC, NCC,
                    # NRAB, Misc, BSUoS) the correct annual cost is:
                    #   total_actual_kWh × annual_mean_rate
                    # We use the price-profile mean rate (not the merged rows), which is
                    # identical for all demand profiles sharing the same tariff year.
                    _ann_kwh_corr = float(active_profile_fin['Overall'].sum() * 0.5)
                    _flat_levy_rate_cols = (
                        [f'{col} [p/kWh] {yr_str}' for col in pl_tariff_cols] +
                        [
                            f'Miscellaneous Charges [p/kWh] {yr_str}',
                            f'BSUoS [p/kWh] {yr_str}',
                        ]
                    )
                    for _rc in _flat_levy_rate_cols:
                        _cc = _rc.replace(f' [p/kWh] {yr_str}', ' Cost [£]')
                        if _rc in year_prices.columns and _cc in ptype_costs_out.columns:
                            ptype_costs_out[_cc] = (
                                _ann_kwh_corr * float(year_prices[_rc].mean()) / 100
                            )
                    # Re-aggregate Policy Levy Costs from the corrected individual columns
                    _pl_ind = [
                        f'{col} Cost [£]' for col in pl_tariff_cols
                        if f'{col} Cost [£]' in ptype_costs_out.columns
                    ]
                    _pl_agg = 'Policy Levy Costs Cost [£]'
                    if _pl_ind and _pl_agg in ptype_costs_out.columns:
                        ptype_costs_out[_pl_agg] = sum(
                            float(ptype_costs_out[c].iloc[0]) for c in _pl_ind
                        )
                    # Re-aggregate Total Cost [£] (if present) from top-level components
                    _total_col_ptype = 'Total Cost [£]'
                    if _total_col_ptype in ptype_costs_out.columns:
                        _top = [
                            'DUoS Volumetric Cost [£]',
                            'Wholesale Commodity Cost [£]',
                            'Policy Levy Costs Cost [£]',
                            'Miscellaneous Charges Cost [£]',
                            'BSUoS Cost [£]',
                            'Capacity Market Cost [£]',
                        ]
                        _top_present = [
                            c for c in _top if c in ptype_costs_out.columns
                        ]
                        if _top_present:
                            ptype_costs_out[_total_col_ptype] = sum(
                                float(ptype_costs_out[c].iloc[0])
                                for c in _top_present
                            )

                ptype_financial = pd.DataFrame()

                # ── DUoS Exceeded Capacity charge ─────────────────────────────────────────
                # Billing-period rule (CDCM / WPD DUoS Charging v6, Table 6): if monthly
                # peak demand exceeds ASC, the excess is charged for ALL days in that month.
                if use_demand_profile and store_temp_out is not None:
                    _hh_kva        = store_temp_out['Overall'] / power_factor
                    _daily_max_kva = _hh_kva.resample('D').max()
                    _monthly_max   = _daily_max_kva.resample('MS').max()
                    _monthly_days  = _daily_max_kva.resample('MS').count()
                    exceed_cap_rate_val = float(
                        add_charges1.loc[
                            (dno_region, yr), 'Exceeded capacity charge [p/kVA/day]'
                        ]
                    )
                    _excess_kva = (_monthly_max - capacity).clip(lower=0)
                    exceeded_cap_cost_dp = float(
                        (_excess_kva * exceed_cap_rate_val * _monthly_days / 100).sum()
                    )

                    # ── Case Study Analytics: per-year metrics ───────────────────
                    _cs_exc = exceeded_cap_incidence(_monthly_max, capacity)
                    _cs_hdr = headroom_utilisation(store_temp_out['Overall'], capacity, power_factor)
                    _cs_row = {
                        'DNO Region':                     dno_region,
                        'Year':                           yr,
                        'Months Breached (Exceeded Cap)': _cs_exc['months_breached'],
                        'Months Total':                   _cs_exc['months_total'],
                        'Proportion Breached [%]':        _cs_exc['proportion_breached'],
                        'Avg Utilisation [% of ASC]':     _cs_hdr['avg_utilisation_pct'],
                        'Peak Utilisation [% of ASC]':    _cs_hdr['peak_utilisation_pct'],
                        'Blended DUoS Rate [p/kWh]':      blended_duos_rate(store_temp_out['Overall'], _hh_rate_srs),
                        'Red-band kWh in TRIAD [%]':      (triad_exposure_share(store_temp_out['Overall'], _hh_band_srs)
                                                           if _hh_band_srs is not None else None),
                    }
                    _cs_key_metrics_rows.append(_cs_row)
                    if _hh_band_srs is not None:
                        _bd = duos_band_breakdown(store_temp_out['Overall'], _hh_band_srs, _hh_rate_srs)
                        _bd.index.name = 'DUoS Band'
                        _bd.insert(0, 'Year', yr)
                        _bd.insert(0, 'DNO Region', dno_region)
                        _cs_band_breakdown_list.append(_bd.reset_index())
                else:
                    exceeded_cap_cost_dp = 0.0

                # ── TNUoS (locational + residual) ──────────────────────────────
                # Locational component: Period 1 LLF × TNUoS locational rate × TRIAD demand
                # Residual component  : non-locational band rate × 365 days (2023+ only)
                if use_demand_profile:
                    dt_series  = active_profile_fin['Date'].dt
                    triad_mask = (
                        (dt_series.weekday < 5) &
                        (dt_series.time >= time(16, 0)) &
                        (dt_series.time < time(19, 0)) &
                        dt_series.month.isin([11, 12, 1, 2])
                    )
                    triad_cons_av = active_profile_fin.loc[
                        triad_mask, 'Overall'
                    ].mean()
                    if pd.isna(triad_cons_av):
                        triad_cons_av = 0
                else:
                    triad_cons_av = current_basis

                tnuos_locational_rate = read_tnuos_locational(
                    dno_no, yr_str, tnuos_locational_df
                )

                try:
                    triad_period_val = (
                        llf_price_reg.loc['Period 1'].iat[0]
                        if not llf_price_reg.empty
                        else 1.0
                    )
                except (KeyError, IndexError):
                    triad_period_val = 1.0

                est_triad_loc       = (
                    triad_period_val * tnuos_locational_rate * triad_cons_av
                )
                const_est_triad_loc = (
                    triad_period_val * tnuos_locational_rate * current_basis
                )

                if int(yr_str) >= 2023:
                    tnuos_nonloc_rate = read_tnuos_nonlocational(
                        current_vol_type, band_number,
                        yr_str, tnuos_nonlocational_df
                    )
                    _days_covered = partial_year_days.get(yr_str) or 365
                    est_res       = tnuos_nonloc_rate * _days_covered
                    est_total     = est_triad_loc + est_res
                    const_est_tot = const_est_triad_loc + est_res

                    if use_flat_load and const_out is not None:
                        const_out[
                            'DUoS Exceeded Capacity Cost [£]'
                        ] = 0.0
                        const_out[
                            'Est. TNUoS Locational Cost [£]'
                        ] = const_est_triad_loc
                        const_out[
                            'Est. TNUoS Residual Cost [£]'
                        ] = est_res
                        const_out[
                            'Est. TNUoS Total Cost [£]'
                        ] = const_est_tot
                        const_out['Year'] = yr
                        const_out = const_out.set_index(
                            'Year', append=True
                        )

                    if use_demand_profile:
                        ptype_financial = ptype_costs_out.copy()
                        ptype_financial[
                            'DUoS Exceeded Capacity Cost [£]'
                        ] = exceeded_cap_cost_dp
                        ptype_financial[
                            'Est. TNUoS Locational Cost [£]'
                        ] = est_triad_loc
                        ptype_financial[
                            'Est. TNUoS Residual Cost [£]'
                        ] = est_res
                        ptype_financial[
                            'Est. TNUoS Total Cost [£]'
                        ] = est_total

                else:
                    if use_flat_load and const_out is not None:
                        const_out[
                            'DUoS Exceeded Capacity Cost [£]'
                        ] = 0.0
                        const_out[
                            'Est. TNUoS Cost [£]'
                        ] = const_est_triad_loc
                        const_out[
                            'Est. TNUoS Total Cost [£]'
                        ] = const_est_triad_loc
                        const_out['Year'] = yr
                        const_out = const_out.set_index(
                            'Year', append=True
                        )

                    if use_demand_profile:
                        ptype_financial = ptype_costs_out.copy()
                        ptype_financial[
                            'DUoS Exceeded Capacity Cost [£]'
                        ] = exceeded_cap_cost_dp
                        ptype_financial[
                            'Est. TNUoS Cost [£]'
                        ] = est_triad_loc
                        ptype_financial[
                            'Est. TNUoS Total Cost [£]'
                        ] = est_triad_loc

                if use_flat_load and const_out is not None:
                    reg_active_fin_list.append(const_out)

                if use_demand_profile:
                    reg_active_fin_list.append(ptype_financial)

            active_fin_full = pd.concat(reg_active_fin_list)
            active_fin_full['Est. TNUoS Total Cost [£]'] = (
                active_fin_full['Est. TNUoS Total Cost [£]']
                .replace(0, np.nan).ffill()
            )
            active_fin_full.index = pd.MultiIndex.from_tuples(
                active_fin_full.index, names=['DNO Region', 'Year']
            )

            active_costs_out = active_fin_full.merge(
                add_charges_out, left_index=True, right_index=True
            )  # Merge annual volumetric and market costs with explicit DUoS fixed, capacity, and reactive annual charges

            active_costs_list.append(active_costs_out)

        analysis_costs_out = pd.concat(active_costs_list)
        peak_split_full    = pd.concat(peak_split_list)
        season_split_full  = pd.concat(season_split_list)

        # ── Assemble quarterly cost DataFrame ─────────────────────────────────
        # Volumetric costs come from HH resampling; fixed annual charges (DUoS
        # Fixed/Capacity/Reactive and TNUoS) are divided by 4 for equal quarterly
        # allocation, then aggregated to the same 6 compressed categories as
        # price_breakdown_agg so the chart reads consistently with other outputs.
        if use_flat_load and quarterly_costs_list:
            _qvol = pd.concat(quarterly_costs_list).reset_index()
            _qvol['Quarter'] = pd.to_datetime(
                _qvol['Quarter Start']
            ).dt.month.map({4: 'Q1', 7: 'Q2', 10: 'Q3', 1: 'Q4'})

            # Divide annual fixed charges by 4 and merge on (DNO Region, Year)
            _annual_fixed = analysis_costs_out[[
                'DUoS Fixed Cost [£]',
                'DUoS Capacity Cost [£]',
                'DUoS Reactive Cost [£]',
                'Est. TNUoS Total Cost [£]',
            ]].copy() / 4
            _annual_fixed = _annual_fixed.reset_index()
            _qvol = _qvol.merge(_annual_fixed, on=['DNO Region', 'Year'], how='left')

            # Aggregate to compressed categories
            _duos_cols = [
                'DUoS Volumetric Cost [£]', 'DUoS Fixed Cost [£]',
                'DUoS Capacity Cost [£]',   'DUoS Reactive Cost [£]',
            ]
            _pl_cols   = [
                'Policy Levy Costs [£]',
                'Capacity Market Cost [£]',
            ]
            _qvol['DUoS [£]']                = _qvol.reindex(columns=_duos_cols).fillna(0).sum(axis=1)
            _qvol['Policy Levy Costs [£]']   = _qvol.reindex(columns=_pl_cols).fillna(0).sum(axis=1)
            # Flag partial quarters: compare actual HH slot count against the
            # minimum expected for a fully-covered quarter of that type.
            # Q1 Apr-Jun=91d, Q2 Jul-Sep=92d, Q3 Oct-Dec=92d, Q4 Jan-Mar=90d.
            _q_min_hh = pd.to_datetime(_qvol['Quarter Start']).dt.month.map(
                {4: 91 * 48, 7: 92 * 48, 10: 92 * 48, 1: 90 * 48}
            )
            _hh_count = (
                _qvol['_hh_count']
                if '_hh_count' in _qvol.columns
                else pd.Series(99999, index=_qvol.index)
            )
            _qvol['Partial'] = _hh_count < _q_min_hh

            quarterly_cost_df = _qvol.rename(columns={
                'Wholesale Commodity Cost [£]':   'Wholesale [£]',
                'BSUoS Cost [£]':                 'BSUoS [£]',
                'Est. TNUoS Total Cost [£]':      'TNUoS [£]',
                'Miscellaneous Charges Cost [£]': 'Miscellaneous [£]',
            })[[
                'DNO Region', 'Year', 'Quarter', 'Partial',
                'Wholesale [£]', 'DUoS [£]', 'BSUoS [£]', 'TNUoS [£]',
                'Policy Levy Costs [£]', 'Miscellaneous [£]',
            ]]
        else:
            quarterly_cost_df = pd.DataFrame()

        if use_demand_profile and len(usecosts_list) > 0:
            use_costs_full = pd.concat(usecosts_list)
        else:
            use_costs_full = pd.DataFrame()

        # ── Step 4: Results export ─────────────────────────────────────────────
        financial_results_dir = f"{analysis_root}/Financial Results"
        os.makedirs(financial_results_dir, exist_ok=True)

        detailed_price_breakdown_df = build_detailed_price_breakdown(
            analysis_costs_out
        )  # Detailed annual price breakdown preserving explicit DUoS and TNUoS components

        detailed_price_breakdown_df = round_export_table(
            detailed_price_breakdown_df, decimals=2
        )  # Rounded detailed annual price breakdown for workbook export

        compressed_price_breakdown_df = build_compressed_price_breakdown(
            analysis_costs_out
        )  # Compressed annual price breakdown combining DUoS and TNUoS into digestible top-level categories

        compressed_price_breakdown_plot = compressed_price_breakdown_df.copy()  # Plot-only compressed bill breakdown table for pie charts
        compressed_price_breakdown_plot['Total Cost [£]'] = compressed_price_breakdown_plot.sum(axis=1)  # Internal annual total used only for pie-chart scaling

        compressed_price_breakdown_df = round_export_table(
            compressed_price_breakdown_df, decimals=2
        )  # Rounded compressed annual price breakdown for workbook export

        save_price_breakdown_workbook(
            compressed_df=compressed_price_breakdown_df,
            detailed_df=detailed_price_breakdown_df,
            workbook_path=f"{financial_results_dir}/Price Type Breakdown.xlsx"
        )

        if use_demand_profile and not use_costs_full.empty:
            use_costs_full.index = pd.MultiIndex.from_tuples(
                use_costs_full.index, names=['DNO Region', 'Year']
            )
            use_costs_full = use_costs_full.rename(
                columns={'Overall Cost [£]': 'Total Cost [£]'}
            )
            use_costs_export = round_export_table(
                reorder_columns_for_export(
                    use_costs_full, demand_breakdown_export_order
                ), decimals=2
            )
            use_costs_export.to_excel(
                f"{financial_results_dir}/Demand Type Breakdown.xlsx"
            )

        for df_save, path in [
            (
                peak_split_full,
                f"{financial_results_dir}/"
                f"Time of Day Split (Peak vs Off-Peak).xlsx"
            ),
            (
                season_split_full,
                f"{financial_results_dir}/"
                f"Seasonal Split (Winter vs Summer).xlsx"
            )
        ]:
            export_df = df_save.copy()
            export_df = export_df.rename(columns={
                c: f'{c} [£]' for c in export_df.columns if '[£]' not in c
            })
            export_df['Total Cost [£]'] = export_df.sum(axis=1)
            export_df.index = pd.MultiIndex.from_tuples(
                export_df.index, names=['DNO Region', 'Year']
            )
            export_df = round_export_table(export_df, decimals=2)
            export_df.to_excel(path)

        if not quarterly_cost_df.empty:
            _q_cost_cols = [c for c in [
                'Wholesale [£]', 'DUoS [£]', 'BSUoS [£]',
                'TNUoS [£]', 'Policy Levy Costs [£]', 'Miscellaneous [£]'
            ] if c in quarterly_cost_df.columns]
            _qdf = quarterly_cost_df.copy()
            _qdf['Total [£]'] = _qdf[_q_cost_cols].sum(axis=1)
            _q_path = f"{financial_results_dir}/Quarterly Cost Breakdown.xlsx"
            with pd.ExcelWriter(_q_path, engine='openpyxl') as _qxl:
                for _qcol in ['Total [£]'] + _q_cost_cols:
                    _pivot = _qdf.pivot_table(
                        index=['DNO Region', 'Year'],
                        columns='Quarter',
                        values=_qcol,
                        aggfunc='sum'
                    ).reindex(columns=['Q1', 'Q2', 'Q3', 'Q4'])
                    _pivot.columns.name = None
                    _pivot.columns = [f'{c} [£]' for c in _pivot.columns]
                    _pivot = round_export_table(_pivot, decimals=2)
                    _pivot.to_excel(_qxl, sheet_name=_qcol.replace(' [£]', '')[:31])

        # ── Step 5: Template-day charts (Generic only, skipped when VAT enabled) ─
        if run_mode == 'Generic' and not apply_vat:
            months = [
                'Jan','Feb','Mar','Apr','May','Jun',
                'Jul','Aug','Sep','Oct','Nov','Dec'
            ]
            month_av_index = [
                f"{mon} ({dt}) [p/kWh]"
                for mon in months
                for dt in ['WD', 'WE']
            ]

            template_day_cache         = {}
            annual_month_profile_cache = {}  # region → {yr_int: DataFrame(48 rows × 12 month cols)}

            for region in run_dno_regions:
                annual_profiles = annual_profiles_cache[region]

                hh_av_month_list = []
                month_av_list    = []

                for yr in list(years_in):
                    yr_str    = str(yr)
                    year_data = annual_profiles[
                        [f'Year {yr_str}', f'Total [p/kWh] {yr_str}']
                    ].dropna().copy()

                    dt_s = pd.to_datetime(
                        year_data[f'Year {yr_str}']
                    ).dt
                    year_data['Month_Num'] = dt_s.month
                    year_data['Time']      = dt_s.time
                    year_data['Day_Type']  = np.where(
                        dt_s.weekday < 5, 'WD', 'WE'
                    )
                    month_map = {i+1: months[i] for i in range(12)}
                    year_data['Month_Str'] = year_data['Month_Num'].map(
                        month_map
                    )
                    year_data['Category']  = (
                        year_data['Month_Str'] + " (" +
                        year_data['Day_Type'] + ") [p/kWh]"
                    )

                    grouped = year_data.groupby(
                        ['Category', 'Time']
                    )[f'Total [p/kWh] {yr_str}'].mean().reset_index()

                    hh_av_month = grouped.pivot(
                        index='Time', columns='Category',
                        values=f'Total [p/kWh] {yr_str}'
                    ).reindex(columns=month_av_index)
                    hh_av_month['Year'] = yr
                    hh_av_month_list.append(
                        hh_av_month.set_index('Year', append=True)
                    )
                    month_avs       = (
                        hh_av_month.drop(columns='Year')
                        .mean().to_frame().T
                    )
                    month_avs.index = [yr]
                    month_av_list.append(month_avs)

                    # Month-only groupby (WD+WE naturally combined by day count)
                    # for Step 6b annual profile charts
                    grp_mo = year_data.groupby(
                        ['Month_Str', 'Time']
                    )[f'Total [p/kWh] {yr_str}'].mean().reset_index()
                    hh_mo = grp_mo.pivot(
                        index='Time', columns='Month_Str',
                        values=f'Total [p/kWh] {yr_str}'
                    ).reindex(columns=months)
                    annual_month_profile_cache.setdefault(region, {})[int(yr_str)] = hh_mo

                hh_av_month_full  = pd.concat(hh_av_month_list)
                month_av_full     = pd.concat(month_av_list)

                hh_av_month_full1 = hh_av_month_full.T
                hh_av_month_full1['DNO Region No.'] = (
                    f'DNO {int(region[:2]):02d}'
                )
                hh_av_month_full2 = (
                    hh_av_month_full1.reset_index()
                    .set_index('DNO Region No.').T.reset_index()
                )

                os.makedirs(
                    f"{chart_root}/Monthly Averages", exist_ok=True
                )

                month_av_full.to_excel(
                    f"{chart_root}/Monthly Averages/{region}"
                )

                # Cache in memory for Step 6 — avoids writing and re-reading
                # Average Template Days from disk (saves one read_excel per region)
                template_day_cache[region] = hh_av_month_full2.reset_index()
                print(f"  ✓ Template Days: {region[:-5]}")

            # ── Step 6: Price-profile line charts (Generic only) ───────────────
            for region in run_dno_regions:
                dno_reg  = region[:-5]
                temp_day = template_day_cache[region]
                data     = reformat(temp_day)
                cmap      = mpl.colormaps['tab10']
                x         = np.arange(0, 48)

                for mon in list(months):
                    mon_short = mon[:3]
                    for day in [' (WD) [p/kWh]', ' (WE) [p/kWh]']:
                        day_short = day[2:4]
                        mon_price = mon_short + day
                        day_data  = data[['Time', 'Year', mon_price]]
                        plot_data = day_data.pivot(
                            index='Time', columns='Year',
                            values=mon_price
                        ).reset_index()

                        fig, ax = plt.subplots()
                        endpoints = []
                        for i, yr_str in enumerate(years_in):
                            yr_int = int(yr_str)
                            color  = cmap(i % 10)
                            y_vals = np.array(plot_data[str(yr_int)])
                            ax.plot(x, y_vals, '-', color=color)
                            endpoints.append((y_vals[-1], str(yr_int), color))

                        all_vals = np.concatenate([
                            np.array(plot_data[str(int(yr))])
                            for yr in years_in
                        ])
                        y_max = np.nanmax(all_vals)
                        ax.set_xlim([0, 53])
                        ax.set_ylim([0, max(y_max * 1.05, 5)])
                        _place_right_labels(ax, endpoints)
                        ax.set_xticks(np.arange(0, 49, 2.0))
                        ax.set_title(
                            f"{mon_short} {day_short}: {dno_reg}\n"
                            f"{current_vol_type} — {current_band}"
                            f"  (excl. fixed, capacity & TNUoS)"
                        )
                        ax.set_xlabel("Settlement Period (30-min, 1–48)")
                        ax.set_ylabel("Volumetric Price Components [p/kWh]")

                        save_dir = f"{chart_root}/Plots/{dno_reg}"
                        os.makedirs(save_dir, exist_ok=True)
                        fig.savefig(
                            f"{save_dir}/{mon} {day_short}.png",
                            dpi=150, bbox_inches='tight'
                        )
                        plt.close()

            # ── Step 6b: Annual profile charts — months as lines (Generic only) ──
            cmap_mo = mpl.colormaps['tab20']
            # Stride-2 indices give maximum colour separation for 12 months
            month_colors = [cmap_mo(i * 2 / 20) for i in range(12)]

            for region in run_dno_regions:
                dno_reg    = region[:-5]
                mo_by_year = annual_month_profile_cache.get(region, {})

                for yr_int, hh_mo in sorted(mo_by_year.items()):
                    x_mo = np.arange(0, 48)
                    fig, ax = plt.subplots()
                    endpoints_mo = []
                    for col_idx, mon in enumerate(months):
                        if mon not in hh_mo.columns:
                            continue
                        color  = month_colors[col_idx]
                        y_vals = hh_mo[mon].values
                        ax.plot(x_mo, y_vals, '-', color=color)
                        endpoints_mo.append((y_vals[-1], mon, color))

                    all_mo_vals = hh_mo.values.flatten()
                    y_max_mo    = np.nanmax(all_mo_vals) if len(all_mo_vals) else 5
                    ax.set_xlim([0, 53])
                    ax.set_ylim([0, max(y_max_mo * 1.05, 5)])
                    _place_right_labels(ax, endpoints_mo)
                    ax.set_xticks(np.arange(0, 49, 2.0))
                    ax.set_title(
                        f"{yr_int}: {dno_reg}\n"
                        f"{current_vol_type} — {current_band}"
                        f"  (excl. fixed, capacity & TNUoS)"
                    )
                    ax.set_xlabel("Settlement Period (30-min, 1–48)")
                    ax.set_ylabel("Volumetric Price Components [p/kWh]")

                    save_dir_mo = f"{chart_root}/Plots/{dno_reg}/Annual Profiles"
                    os.makedirs(save_dir_mo, exist_ok=True)
                    fig.savefig(
                        f"{save_dir_mo}/{yr_int}.png",
                        dpi=150, bbox_inches='tight'
                    )
                    plt.close()

                print(f"  ✓ Annual Profiles: {dno_reg}")

        # ── Step 7: Financial summary charts (both modes) ─────────────────────
        
        analysis_costs_prepped = prep_costs(
            analysis_costs_out
        )  # Financial output table used for charts and regional comparisons with explicit DUoS component columns preserved

        if apply_vat:
            _vat_col = analysis_costs_prepped['Total Cost [£]'] * vat_rate
            analysis_costs_prepped['VAT [£]'] = _vat_col
            analysis_costs_prepped['Total Cost [£]'] = (
                analysis_costs_prepped['Total Cost [£]'] + _vat_col
            )  # VAT added as named cost category; Total Cost [£] updated to VAT-inclusive for ECB and all downstream charts

        if run_mode == 'Generic':
            annual_energy_kwh = current_basis * 8760  # Annual Generic flat-load energy demand used to convert annual £ cost into effective p/kWh

            annual_p_per_kwh_df = build_effective_cost_benchmark(
                analysis_costs_prepped=analysis_costs_prepped,
                annual_energy_kwh=annual_energy_kwh
            )  # Regional annual effective electricity cost benchmark for the current Generic voltage and band run

            annual_p_per_kwh_df = round_export_table(
                annual_p_per_kwh_df, decimals=2
            )  # Rounded annual effective electricity cost benchmark for workbook export

            ecb_collection[(current_vol_type, current_band)] = annual_p_per_kwh_df

            effective_cost_workbook_path = (
                f"{analysis_root}/Effective Cost Benchmark.xlsx"
            )  # Output workbook for the Generic flat-load effective electricity cost benchmark

            save_effective_cost_benchmark(
                annual_p_per_kwh_df=annual_p_per_kwh_df,
                workbook_path=effective_cost_workbook_path
            )

        elif use_demand_profile and _specific_annual_kwh is not None:
            annual_energy_kwh = _specific_annual_kwh  # Actual annual kWh from the demand profile (kW × 0.5 h, summed)

            annual_p_per_kwh_df = build_effective_cost_benchmark(
                analysis_costs_prepped=analysis_costs_prepped,
                annual_energy_kwh=annual_energy_kwh
            )  # Specific-mode effective electricity cost benchmark

            annual_p_per_kwh_df = round_export_table(
                annual_p_per_kwh_df, decimals=2
            )

            effective_cost_workbook_path = (
                f"{analysis_root}/Effective Cost Benchmark.xlsx"
            )
            annual_p_per_kwh_df.reset_index().to_excel(
                effective_cost_workbook_path, index=False
            )

        compressed_price_breakdown_plot = build_compressed_price_breakdown(
            analysis_costs_prepped
        )  # Compressed annual bill breakdown table used for clearer price-type pie charts

        compressed_price_breakdown_plot['Total Cost [£]'] = (
            compressed_price_breakdown_plot.sum(axis=1)
        )  # Internal annual total used only for pie-chart percentage scaling

        if apply_vat:
            _cbp_vat = compressed_price_breakdown_plot['Total Cost [£]'] * vat_rate
            compressed_price_breakdown_plot['VAT [£]'] = _cbp_vat
            compressed_price_breakdown_plot['Total Cost [£]'] = (
                compressed_price_breakdown_plot['Total Cost [£]'] + _cbp_vat
            )

        # Exclude partial years from pie and stacked-line charts
        if _partial_yr_set:
            compressed_price_breakdown_plot = compressed_price_breakdown_plot[
                ~compressed_price_breakdown_plot.index
                .get_level_values('Year').astype(str)
                .isin(_partial_yr_set)
            ]

        summary_tables = create_summary_tables(
            analysis_costs_prepped=analysis_costs_prepped,
            peak_split_full=peak_split_full,
            season_split_full=season_split_full,
            use_costs_full=use_costs_full,
            use_demand_profile=use_demand_profile,
            partial_yr_set=_partial_yr_set
        )

        if use_demand_profile and _cs_band_breakdown_list:
            summary_tables['duos_band_breakdown'] = pd.concat(
                _cs_band_breakdown_list, ignore_index=True
            )
        if use_demand_profile and _cs_key_metrics_rows:
            summary_tables['case_study_metrics'] = pd.DataFrame(_cs_key_metrics_rows)

        for key, preferred_order in [
            ('average_annual_costs', price_breakdown_export_order),
            (
                'price_type_breakdown_detailed',
                ['DNO Region', 'Year'] + (
                    [
                        'DUoS Volumetric Cost [£]',
                        'DUoS Fixed Cost [£]',
                        'DUoS Capacity Cost [£]',
                        'DUoS Reactive Cost [£]',
                        'DUoS Exceeded Capacity Cost [£]',
                        'Wholesale Commodity Cost [£]',
                        'Policy Levy Costs [£]'
                    ] +
                    pl_cost_cols +
                    [
                        'Miscellaneous Charges Cost [£]'
                    ] +
                    misc_cost_cols +
                    [
                        'BSUoS Cost [£]',
                        'Capacity Market Cost [£]',
                        'Est. TNUoS Cost [£]',
                        'Est. TNUoS Locational Cost [£]',
                        'Est. TNUoS Residual Cost [£]',
                        'Total Cost incl. VAT [£]'
                    ]
                )
            ),
            (
                'price_type_breakdown_compressed',
                ['DNO Region', 'Year'] + [
                    'DUoS Total Cost [£]',
                    'Wholesale Commodity Cost [£]',
                    'Policy Levy Costs [£]',
                    'Miscellaneous Charges Cost [£]',
                    'BSUoS Cost [£]',
                    'Capacity Market Cost [£]',
                    'TNUoS Total Cost [£]'
                ]
            ),
            (
                'demand_type_breakdown',
                ['DNO Region', 'Year'] + demand_breakdown_export_order
            )
        ]:
            if key in summary_tables:
                summary_tables[key] = round_export_table(
                    reorder_columns_for_export(
                        summary_tables[key], preferred_order
                    ), decimals=2
                )

        for key in [
            'annual_total_costs', 'seasonal_split',
            'peak_vs_off_peak', 'key_metrics'
        ]:
            if key in summary_tables:
                summary_tables[key] = round_export_table(
                    summary_tables[key], decimals=2
                )

        summary_workbook_path = f"{analysis_root}/Summary Workbook.xlsx"
        save_summary_workbook(
            summary_tables=summary_tables,
            summary_workbook_path=summary_workbook_path
        )

        for region in run_dno_regions:
            dno_reg = region[:-5]

            pie_chart(
                compressed_price_breakdown_plot,
                'Price Type Breakdown',
                analysis_label,
                dno_reg,
                current_band,
                current_vol_type,
                analysis_root
            )  # Annual price-type pie charts based on compressed top-level bill categories

            bill_breakdown_stacked_line(
                compressed_price_breakdown_plot, 'Price Type Breakdown',
                analysis_label, dno_reg, current_band, current_vol_type, analysis_root
            )  # Stacked-area supplement showing all years of the price-type breakdown together

            if use_demand_profile and not use_costs_full.empty:
                pie_chart(
                    use_costs_full,
                    'Usage Type Breakdown',
                    analysis_label,
                    dno_reg,
                    current_band,
                    current_vol_type,
                    analysis_root
                )  # Annual demand-type pie charts for demand-profile runs only

        ordered_dno = analysis_costs_prepped.xs(
            years_in[1], level='Year'
        ).sort_values('Total Cost [£]').reset_index()

        _n            = len(ordered_dno)
        cmap_oranges  = mpl.colormaps['Oranges']
        hex_dict_full = {
            str(r): mpl.colors.rgb2hex(
                cmap_oranges(0.35 + 0.65 * i / max(_n - 1, 1))
            )
            for i, r in enumerate(ordered_dno['DNO Region'])
        }

        peak_split_plot           = peak_split_full.copy()
        peak_split_plot['Total Cost [£]'] = peak_split_plot.sum(axis=1)
        season_split_plot         = season_split_full.copy()
        season_split_plot['Total Cost [£]'] = (
            season_split_plot.sum(axis=1)
        )

        if apply_vat:
            peak_split_plot   = peak_split_plot   * (1 + vat_rate)
            season_split_plot = season_split_plot * (1 + vat_rate)

        price_breakdown_agg = build_compressed_price_breakdown(analysis_costs_prepped).copy()  # Top-level aggregated Price Breakdown view (matches pie_chart categories)
        if 'Capacity Market Cost [£]' in price_breakdown_agg.columns:
            price_breakdown_agg['Policy Levy Costs [£]'] += (
                price_breakdown_agg.pop('Capacity Market Cost [£]')
            )
        price_breakdown_agg = price_breakdown_agg.rename(columns={
            'Policy Levy Costs [£]': 'Policy Levy Costs [£]',
            'Miscellaneous Charges Cost [£]':              'Miscellaneous [£]',
            'DUoS Total Cost [£]':                         'DUoS [£]',
            'Wholesale Commodity Cost [£]':                'Wholesale [£]',
            'BSUoS Cost [£]':                              'BSUoS [£]',
            'TNUoS Total Cost [£]':                        'TNUoS [£]',
        })
        price_breakdown_agg['Total Cost [£]'] = price_breakdown_agg.sum(axis=1)
        if apply_vat:
            _agg_vat = price_breakdown_agg['Total Cost [£]'] * vat_rate
            price_breakdown_agg['VAT [£]'] = _agg_vat
            price_breakdown_agg['Total Cost [£]'] = (
                price_breakdown_agg['Total Cost [£]'] + _agg_vat
            )
        price_breakdown_agg = reorder_columns_for_export(
            price_breakdown_agg,
            ['Wholesale [£]', 'DUoS [£]', 'BSUoS [£]', 'TNUoS [£]',
             'Policy Levy Costs [£]', 'Miscellaneous [£]', 'VAT [£]']
        )  # Wholesale, then Network (DUoS/BSUoS/TNUoS sequential), then Policy Levy, then Misc, then VAT

        _disagg_drop = [
            c for c in analysis_costs_prepped.columns
            if 'Policy Levy Costs' in c
            or 'Miscellaneous Charges Cost' in c
            or c == 'Est. TNUoS Total Cost [£]'  # exclude aggregate; sub-components (Est. TNUoS Cost/Locational/Residual) are already present
        ]
        price_breakdown_disagg = (
            analysis_costs_prepped
            .drop(columns=_disagg_drop, errors='ignore')
            .rename(columns=PRICE_LABEL_MAP)
        )  # Full disaggregated Price Breakdown view with long-form component labels
        price_breakdown_disagg = reorder_columns_for_export(
            price_breakdown_disagg, PRICE_BREAKDOWN_DISAGG_ORDER
        )

        plot_dict_active = {
            'Colour Visualisation':          analysis_costs_prepped,
            'Price Breakdown Aggregated':     price_breakdown_agg,
            'Price Breakdown Disaggregated':  price_breakdown_disagg,
            'Winter Vs Summer':              season_split_plot,
            f'Peak ({peak_label}) Vs NPeak': peak_split_plot
        }

        if use_demand_profile and not use_costs_full.empty:
            use_costs_plot = use_costs_full.copy()
            if (
                'Total Cost [£]' not in use_costs_plot.columns and
                'Overall Cost [£]' in use_costs_plot.columns
            ):
                use_costs_plot = use_costs_plot.rename(
                    columns={'Overall Cost [£]': 'Total Cost [£]'}
                )
            plot_dict_active['Demand Breakdown'] = use_costs_plot

        _national_trend_labels = {
            'Price Breakdown Aggregated',
            'Price Breakdown Disaggregated',
            f'Peak ({peak_label}) Vs NPeak',
        }

        # Exclude partial years from annual cost plots — a part-year total looks
        # artificially small because demand for the full year has not yet occurred.
        if _partial_yr_set:
            plot_dict_active = {
                lbl: df[
                    ~df.index.get_level_values('Year').astype(str).isin(_partial_yr_set)
                ]
                for lbl, df in plot_dict_active.items()
            }

        for label, input_data in plot_dict_active.items():
            stacked_plot(input_data,     label, analysis_label, current_band, current_vol_type, analysis_root, hex_dict_full)
            avs_stacked_plot(input_data, label, analysis_label, current_band, current_vol_type, analysis_root, hex_dict_full)
            if label == 'Colour Visualisation':
                colour_vis_trend_plot(input_data, analysis_label, current_band, current_vol_type, analysis_root, hex_dict_full)
            if label == 'Winter Vs Summer':
                season_trend_plot(input_data, analysis_label, current_band, current_vol_type, analysis_root)
            if label in _national_trend_labels:
                national_trend_plot(input_data, label, analysis_label, current_band, current_vol_type, analysis_root,
                                    footnote_text=_footnote_annual)

        if not quarterly_cost_df.empty:
            _qdf_plot = quarterly_cost_df.copy()
            if apply_vat:
                _q_cost_cols = [c for c in _qdf_plot.columns
                                if c not in ('DNO Region', 'Year', 'Quarter', 'Partial')]
                _qdf_plot['VAT [£]'] = _qdf_plot[_q_cost_cols].sum(axis=1) * vat_rate
            quarterly_cost_plot(
                _qdf_plot, analysis_label, current_band,
                current_vol_type, analysis_root
            )

        print(
            f"\n  ✓ Completed: {current_vol_type} — {current_band}"
        )

if run_mode == 'Generic' and ecb_collection:
    cross_vol_dir = (
        f"{output_root}/Cross Voltage Comparison (incl. VAT)"
        if apply_vat else
        f"{output_root}/Cross Voltage Comparison"
    )
    os.makedirs(cross_vol_dir, exist_ok=True)
    ecb_comp_path = f"{cross_vol_dir}/Effective Cost Benchmark Comparison.xlsx"

    ordered_keys = [
        (vt, band)
        for vt in all_vol_types
        for band in ['Band 1', 'Band 2', 'Band 3', 'Band 4']
        if (vt, band) in ecb_collection
    ]

    with pd.ExcelWriter(ecb_comp_path, engine='openpyxl') as writer:
        # Summary sheet first: rows = vol_type | band (regions aggregated), columns = years
        summary_rows = {}
        for (vt, band) in ordered_keys:
            df = ecb_collection[(vt, band)]  # index=regions, cols=years
            vt_short = 'ND Aggregated' if vt == 'Non-Domestic Aggregated' else vt
            summary_rows[f"{vt_short} | {band}"] = df.mean(axis=0)  # average across regions
        summary_df = pd.DataFrame(summary_rows).T  # rows=labels, cols=years
        summary_df.round(2).to_excel(writer, sheet_name='Summary')

        for (vt, band) in ordered_keys:
            vt_short = 'ND Aggregated' if vt == 'Non-Domestic Aggregated' else vt
            sheet_name = f"{vt_short} {band}"
            ecb_collection[(vt, band)].to_excel(writer, sheet_name=sheet_name)

    plot_ecb_trend(ecb_collection, all_vol_types, cross_vol_dir,
                   partial_yr_set=_partial_yr_set)
    print(f"  ✓ Cross-voltage ECB comparison saved: {ecb_comp_path}")

print("\n--- All Processing Completed Successfully! ---")

# --- END TIMER ---
end_time               = pytime.perf_counter()
execution_time_seconds = end_time - start_time
minutes = int(execution_time_seconds // 60)
seconds = execution_time_seconds % 60
print(
    f"\nTotal programme run time: "
    f"{minutes} minutes and {seconds:.2f} seconds"
)