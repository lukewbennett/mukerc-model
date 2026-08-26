#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Build Price Profiles v8
=======================
Script 3 of 4 in the MUKERC pipeline.

Generates half-hourly (HH) electricity price profiles for every combination
of voltage type, residual charging band, and DNO region. Profiles are saved to:
    Output Data Repository/Generic Results/{vol_type}/{band}/
        1 - Price Profiles/Half-Hourly Prices/

Price components written to each profile:
  - DUoS Actual (Green/Amber/Red volumetric charges)
  - Policy Levy Costs and sub-components
  - Miscellaneous Charges and sub-components
  - BSUoS (seasonal, summer/winter)
  - Wholesale Commodity (Elexon HH mid prices)
  - Capacity Market (weekday 16:00–19:00 window only)

Run behaviour is controlled by Manual_Input.py:
  - build_scope_mode  : 'Full' (all combinations) or 'Test' (restricted subset)
  - start_year        : first financial year (April start) to include
  - end_year_limit    : last financial year to include

This script must be run before Post Analysis.py.

Created on Sun Jun 07 2026
@author: Luke Bennett
"""

import os
import pandas as pd
import numpy as np
from datetime import time
from matplotlib import pyplot as plt
import matplotlib as mpl
plt.style.use('seaborn-v0_8-whitegrid')
import re
import time as pytime

# --- START TIMER ---
start_time = pytime.perf_counter()

# =============================================================================
# SECTION 1: SELECT INPUT PARAMETERS
# =============================================================================

from Manual_Input import (
    all_dno_regions, all_vol_types,
    start_year, end_year_limit,
    build_scope_mode, build_test_dno_regions,
    build_test_vol_types, build_test_bands
)

from Modules.Helper_Functions import (
    get_vol_row_label, get_annex_sheet,
    prepare_tariff_tables, prepare_seasonal_component_tables,
    load_bsuos_rates
)

# NOTE: Build Price Profiles.py now runs in Generic mode only.
# It generates price profiles for all voltage types, all bands
# and all DNO regions. These are then read by Post Analysis.py
# in both Generic and Specific modes.
#
# price_type is hardcoded as 'Import'.
# Export functionality has been removed from the MUKERC framework.

price_type      = 'Import'
run_vol_types   = all_vol_types
run_dno_regions = all_dno_regions
output_root     = "../Output Data Repository/Generic Results"
bands_to_run    = ['Band 1', 'Band 2', 'Band 3', 'Band 4']

if build_scope_mode == 'Test':
    run_vol_types   = build_test_vol_types
    run_dno_regions = build_test_dno_regions
    bands_to_run    = build_test_bands

print(
    f"\n✓ Build Price Profiles — Generic Mode "
    f"({build_scope_mode} Scope)\n"
    f"  {len(run_vol_types)} voltage types × "
    f"{len(run_dno_regions)} DNO regions × "
    f"{len(bands_to_run)} bands"
)

# =============================================================================
# FORMAT INPUT PARAMETERS
# =============================================================================

years_past_all = pd.Series(
    os.listdir("Input/Network Prices/Distribution")
).replace(".DS_Store", np.nan).dropna().sort_values()

if years_past_all.empty:
    raise ValueError("No distribution input year folders were found.")

years_past = years_past_all[
    years_past_all.astype(int).between(start_year, end_year_limit)
].sort_values()

if years_past.empty:
    raise ValueError(
        f"No distribution input years found between "
        f"start_year={start_year} and end_year_limit={end_year_limit}."
    )

build_end_year     = int(years_past.astype(int).max())
build_end_year_str = str(build_end_year)

print(
    f"  Years: {start_year} → {build_end_year} "
    f"({len(years_past)} financial years)"
)

# =============================================================================
# FETCH WHOLESALE, ENVIRONMENTAL AND CAPACITY-MARKET INPUT DATA
# =============================================================================

tariff_sum = pd.read_excel(
    "Input/Policy Levy Prices/Policy Levy Rates.xlsx",
    header=3
)

misc_charges = pd.read_excel(
    "Input/Miscellaneous Prices/Miscellaneous Rates.xlsx",
    header=3
)  # Seasonal miscellaneous charge workbook with the same year/season structure as the Policy Levy rates workbook

elexon_prices_raw = pd.read_excel(
    "Input/Commodity Prices/compiled_elexon_mid_prices.xlsx"
)  # Raw half-hourly Elexon mid-price table used to replace legacy N2EX block prices

pl_tariff_cols, pl_tariff_table, pl_tariff_total, cm_table = (
    prepare_tariff_tables(tariff_sum)
)  # Dynamic tariff workbook interpretation: all non-id, non-cm columns are treated as Policy Levy tariff components

misc_charge_cols, misc_charge_table, misc_charge_total = (
    prepare_seasonal_component_tables(misc_charges)
)  # Dynamic miscellaneous seasonal charge interpretation: all non-id columns are treated as miscellaneous charge components

bsuos_rates = load_bsuos_rates(
    "Input/Network Prices/Balancing/BSUoS Rates.xlsx"
)  # Seasonal BSUoS rates in £/MWh; divided by 10 at point of use to convert to p/kWh

# =============================================================================
# STANDALONE HELPER FUNCTIONS
# =============================================================================

def replace_time(x):
    """Strips non-numeric characters and normalises midnight to 2330."""
    x = re.sub("[^0-9]", "", x)
    x = x.replace('\n', '').replace('2400', '2330')
    return x


def split_times(time_tab, labels, col, days, t3):
    """
    Populates time_tab with 4-character time strings (HHMM) for each
    Green/Amber/Red band label across Weekday and Weekend rows.

    Parameters
    ----------
    time_tab : DataFrame  index=band labels (Gr1..Re6), columns=['Weekday','Weekend']
    labels   : Series     band labels to fill (e.g. gr_labels = ['Gr1'..'Gr6'])
    col      : str        colour column to read from t3 ('Green', 'Amber', 'Red')
    days     : Series     ['Weekday', 'Weekend']
    t3       : DataFrame  2-row table (Weekday / Weekend) of concatenated HHMM strings
    """
    for day in list(days):
        ind_time = t3.loc[day, col]
        x, x1 = -4, 0
        for time_per in list(labels):
            x, x1 = x + 4, x1 + 4  # advance by 4 characters to extract next HHMM block
            ind_time1 = ind_time[x:x1]
            if ind_time1 == '':
                ind_time1 = np.nan
            time_tab.loc[time_per, day] = ind_time1
    return time_tab


def load_elexon_hh_prices(elexon_prices_raw_df):
    """
    Loads and validates pre-cleaned half-hourly Elexon mid prices.

    Required columns:
    - startTime
    - combined_price

    Assumptions:
    - startTime is already a cleaned naive half-hour timestamp
      aligned to the model timeline
    - combined_price is in £/MWh
    - output is converted to p/kWh
    """
    required_cols = ['startTime', 'combined_price']
    missing_cols = [
        col for col in required_cols if col not in elexon_prices_raw_df.columns
    ]
    if missing_cols:
        raise ValueError(
            f"Elexon price input is missing required column(s): {missing_cols}. "
            f"Required columns are: {required_cols}."
        )

    elexon_prices = elexon_prices_raw_df[
        ['startTime', 'combined_price']
    ].copy()  # Pre-cleaned half-hour timestamp and Elexon market price columns used to build the wholesale commodity series

    elexon_prices['DateTime'] = pd.to_datetime(
        elexon_prices['startTime'], utc=True
    ).dt.tz_localize(None)  # Parse the cleaned half-hour timestamps and remove UTC timezone awareness so they align with the model's naive datetime timeline

    elexon_prices['combined_price'] = pd.to_numeric(
        elexon_prices['combined_price'], errors='coerce'
    )  # Convert pre-cleaned Elexon prices to numeric values for validation and unit conversion

    if elexon_prices['DateTime'].duplicated().any():
        raise ValueError(
            "Elexon price input contains duplicate half-hour timestamps in "
            "'startTime' after pre-cleaning."
        )

    if elexon_prices['combined_price'].isna().any():
        raise ValueError(
            "Elexon price input contains non-numeric or missing values in "
            "'combined_price' after pre-cleaning."
        )

    elexon_prices['Wholesale Commodity [p/kWh]'] = (
        elexon_prices['combined_price'] / 10
    )  # Convert Elexon market prices from £/MWh to p/kWh for direct use in half-hourly profile construction

    elexon_prices = elexon_prices[[
        'DateTime', 'Wholesale Commodity [p/kWh]'
    ]].sort_values('DateTime')  # Final cleaned half-hourly wholesale commodity price table used in the Generic profile build

    return elexon_prices


def fetch_prices(vol_type, dno_regions_list, band):
    """
    Fetches historical DNO DUoS tariff prices (Green, Amber, Red)
    for a given voltage level, band and list of regions.

    Parameters
    ----------
    vol_type         : voltage level string
    dno_regions_list : list of DNO region filenames
    band             : residual charging band string e.g. 'Band 4'
    """
    tariff_list = []

    for year in years_past:
        vol     = get_vol_row_label(vol_type, year, band)
        annex_1 = get_annex_sheet(year)

        for dno_region in dno_regions_list:
            dno_reg = dno_region[:-5]
            try:
                df = pd.read_excel(
                    f"Input/Network Prices/Distribution/"
                    f"{year}/{dno_region}",
                    sheet_name=annex_1
                )
            except Exception as e:
                raise RuntimeError(
                    f"Cannot read DNO file '{dno_region}' ({year}) — "
                    f"check the file is not open in Excel. Detail: {e}"
                ) from e

            d1 = df.set_index(df.columns[0])
            d1.columns = d1.loc['Tariff name']

            d3 = d1.loc[vol].reset_index()
            d3['Tariff name'] = d3['Tariff name'].str.lower()

            green = d3[d3['Tariff name'].str.contains("green")].iat[0, 1]
            amber = d3[d3['Tariff name'].str.contains("amber")].iat[0, 1]
            red   = d3[d3['Tariff name'].str.contains("red")].iat[0, 1]

            tariffs = pd.DataFrame(
                [[dno_reg, year, green, amber, red]],
                columns=[
                    'DNO Region', 'Year',
                    'Green [p/kWh]', 'Amber [p/kWh]', 'Red [p/kWh]'
                ]
            )
            tariffs[
                ['Green [p/kWh]', 'Amber [p/kWh]', 'Red [p/kWh]']
            ] = tariffs[
                ['Green [p/kWh]', 'Amber [p/kWh]', 'Red [p/kWh]']
            ].abs()
            tariff_list.append(tariffs)

    past_pr  = pd.concat(tariff_list, ignore_index=True)
    past_pr1 = past_pr.set_index(
        ['DNO Region', 'Year']
    ).sort_index().reset_index().set_index('DNO Region')
    return past_pr1


def fetch_time_periods(dno_regions_list):
    """
    Reads DNO Annex 1 time-band definitions (Green/Amber/Red windows)
    for every year and region, returning a consolidated reference table.

    Parameters
    ----------
    dno_regions_list : list of DNO region filenames (e.g. '10 - Eastern.xlsx')

    Returns
    -------
    DataFrame with MultiIndex (DNO Region, Year), two rows per entry
    (Weekday / Weekend), columns Gr1..Re6 containing 4-character HHMM
    time strings for each colour-band window.
    """
    gr_labels    = pd.Series(['Gr1', 'Gr2', 'Gr3', 'Gr4', 'Gr5', 'Gr6'])
    am_labels    = pd.Series(['Am1', 'Am2', 'Am3', 'Am4', 'Am5', 'Am6'])
    re_labels    = pd.Series(['Re1', 'Re2', 'Re3', 'Re4', 'Re5', 'Re6'])
    all_labels   = pd.concat([gr_labels, am_labels, re_labels])
    full_columns = pd.concat([pd.Series(['Weekday or Weekend']), all_labels])
    time_pers_full = pd.DataFrame(columns=full_columns)

    for year in list(years_past):
        annex_1   = get_annex_sheet(year)
        time_pers = pd.DataFrame(columns=full_columns)

        for dno_region in dno_regions_list:
            dno_reg = dno_region[:-5]
            try:
                df = pd.read_excel(
                    f"Input/Network Prices/Distribution/"
                    f"{year}/{dno_region}",
                    sheet_name=annex_1
                )
            except Exception as e:
                raise RuntimeError(
                    f"Cannot read time periods from '{dno_region}' ({year}) — "
                    f"check the file is not open in Excel. Detail: {e}"
                ) from e

            t         = df.set_index(df.columns[0])
            t.columns = t.loc['Time periods']
            t1 = t[
                ['Red Time Band', 'Amber Time Band', 'Green Time Band']
            ].iloc[:, :3]
            t1['Green Time Band'] = t1['Green Time Band'].bfill()
            t2 = t1[t1['Green Time Band'].str.contains('00', na=False)]
            if len(t2) > 2:
                t2.iloc[:3, :] = t2.iloc[:3, :].ffill().bfill().values
                t2 = t2.iloc[2:, :]
            t2 = t2.fillna('').replace(0, '')

            colour = pd.Series(['Green', 'Amber', 'Red'])
            for col in list(colour):
                t2[col] = t2[col + ' Time Band'].apply(
                    lambda x: replace_time(x)
                )

            t3   = t2.iloc[:, 3:]
            days = pd.Series(['Weekday', 'Weekend'])
            t3   = t3.set_index(days)
            time_tab = pd.DataFrame(index=all_labels, columns=days)

            time_tab = split_times(time_tab, gr_labels, 'Green', days, t3)
            time_tab = split_times(time_tab, am_labels, 'Amber', days, t3)
            time_tab = split_times(
                time_tab, re_labels, 'Red', days, t3
            ).T

            time_tab['DNO Region'] = dno_reg
            y1 = time_tab.reset_index().set_index('DNO Region').rename(
                columns={'index': 'Weekday or Weekend'}
            )
            y1['Year'] = year
            y2 = y1.set_index('Year', append=True)
            time_pers = pd.concat([time_pers, y2])

        time_pers_full = pd.concat([time_pers_full, time_pers])

    time_pers_full.index = pd.MultiIndex.from_tuples(
        time_pers_full.index, names=['DNO Region', 'Year']
    )
    return time_pers_full.sort_index(level='DNO Region')


def plot_dno_prices(past_pr, voltage, out_path, dno_regions_list):
    """
    Saves a line chart of historical Green/Amber/Red DUoS volumetric rates
    for each DNO region to out_path/DUoS Volumetric Charges/.

    Parameters
    ----------
    past_pr          : DataFrame   output of fetch_prices() — DUoS tariff table
                                   indexed by DNO Region
    voltage          : str         voltage type label used in chart title
    out_path         : str         directory to write chart PNGs into;
                                   a 'DUoS Volumetric Charges/' sub-folder is
                                   created automatically
    dno_regions_list : list        DNO region filenames (used to get short labels)
    """
    for dno_reg in dno_regions_list:
        dno_region = dno_reg[:-5]  # Short DNO region label used in chart titles and filenames
        u = past_pr.loc[dno_region]  # Historical DUoS volumetric tariff table for one DNO region
        all_years = u['Year'].astype(float).values  # All tariff years available for this region and voltage type

        tariff_types = ['Green [p/kWh]', 'Amber [p/kWh]', 'Red [p/kWh]']  # DUoS volumetric tariff colour bands
        colours = ['g', 'gold', 'r']  # Plot colours used for Green, Amber and Red lines

        plt.figure(figsize=(10, 6))

        bw      = 0.2
        offsets = [-0.25, 0.0, 0.25]  # centres for Green, Amber, Red bars

        fig, ax = plt.subplots(figsize=(12, 6))
        for tarr_type, col, offset in zip(tariff_types, colours, offsets):
            y_all = np.array(u[tarr_type].astype(float))
            ax.bar(all_years + offset, y_all, width=bw, color=col,
                   label=tarr_type, zorder=3)
            zero_mask = (y_all == 0)
            if zero_mask.any():
                ax.plot(all_years[zero_mask] + offset, np.zeros(zero_mask.sum()),
                        'o', color=col, ms=5, zorder=4)

        ax.set_xlim(all_years[0] - 0.6, all_years[-1] + 0.6)
        ax.set_xticks(all_years)
        ax.set_xticklabels([str(int(y)) for y in all_years])
        ax.yaxis.set_major_formatter(
            mpl.ticker.FuncFormatter(lambda x, p: f'{x:,.0f}')
        )
        ax.set_xlabel('Financial Year')
        ax.set_ylabel('Import Price [p/kWh]')
        ax.set_title(f'Import: {dno_region}\n{voltage}')
        ax.legend()

        save_dir = f"{out_path}/DUoS Volumetric Charges/"
        os.makedirs(save_dir, exist_ok=True)
        fig.savefig(f"{save_dir}{dno_region}.png", dpi=150, bbox_inches='tight')
        plt.close(fig)


def apply_bands(df, duos_times, day_type, duos_name, times_use1,
                prices_actual):
    """
    Stamps DUoS Green/Amber/Red rates onto the correct half-hourly rows,
    and writes the band label ('Green', 'Amber', or 'Red') to a companion
    column named 'DUoS Band {year}' (derived by replacing 'DUoS Actual [p/kWh]'
    in duos_name).

    Parameters
    ----------
    df            : DataFrame   year base frame with 'Time' and 'Weekday' columns
    duos_times    : DataFrame   category table (wd_duos_times or we_duos_times)
                                with 'Categories' and 'Price Category' columns
    day_type      : str         'Weekday' or 'Weekend'
    duos_name     : str         output column name, e.g. 'DUoS Actual [p/kWh] 2022'
    times_use1    : DataFrame   region × year slice of time_pers_final
    prices_actual : DataFrame   single-row DUoS tariff for the year being built
    """
    band_col = duos_name.replace('DUoS Actual [p/kWh]', 'DUoS Band')

    day_mask = (
        df['Weekday'] < 5 if day_type == 'Weekday' else df['Weekday'] >= 5
    )  # 0–4 = Mon–Fri, 5–6 = Sat–Sun

    times_use2 = times_use1.loc[
        times_use1['Weekday or Weekend'] == day_type
    ].T.fillna('0000').iloc[1:]
    times_use2.iloc[:, 0] = pd.to_datetime(
        times_use2.iloc[:, 0], format='%H%M'
    ).dt.time

    # [:-1][::2] selects band-start labels; [1:][::2] selects matching band-end labels
    # — the category list stores alternating (start, end) pairs e.g. [Gr1, Gr2, Am1, Am2, ...]
    for (duos_time, duos_time1, price_cat) in zip(
        duos_times['Categories'][:-1][::2],
        duos_times['Categories'][1:][::2],
        duos_times['Price Category'][::2]
    ):
        start_t   = times_use2.loc[duos_time].iat[0]
        end_t     = times_use2.loc[duos_time1].iat[0]
        time_mask = (df['Time'] >= start_t) & (df['Time'] < end_t)

        df.loc[day_mask & time_mask, duos_name] = \
            prices_actual[price_cat + " [p/kWh]"].iat[0]
        df.loc[day_mask & time_mask, band_col] = price_cat

    return df


def voltage_apply(dno_prices_full, base_yr, curr_yr_out,
                  current_vol_type, y, dno_region, reference_year,
                  time_pers_final, wd_duos_times, we_duos_times):
    """
    Assigns DUoS Actual band prices to every half-hourly row for one DNO
    region and year, merging the result into the cumulative output frame.

    Parameters
    ----------
    dno_prices_full  : DataFrame   full DUoS tariff table (output of fetch_prices)
    base_yr          : DataFrame   pre-built year base frame (from year_base_frames)
    curr_yr_out      : DataFrame   cumulative output frame for this region
    current_vol_type : str         voltage type (informational, not used in logic)
    y                : str         financial year string e.g. '2022'
    dno_region       : str         short DNO region label (no .xlsx extension)
    reference_year   : str         year whose time-band windows are used for DUoS
                                   assignment; currently always build_end_year_str
    time_pers_final  : DataFrame   output of fetch_time_periods()
    wd_duos_times    : DataFrame   weekday band category table
    we_duos_times    : DataFrame   weekend band category table

    Returns
    -------
    curr_yr_out : DataFrame updated with 'DUoS Actual [p/kWh] {y}' and
                  'DUoS Band {y}' columns added
    """
    df            = base_yr.copy()
    duos_name     = f"DUoS Actual [p/kWh] {y}"
    band_col      = f"DUoS Band {y}"
    df[duos_name] = np.nan
    df[band_col]  = ''

    past_pr1      = dno_prices_full.loc[dno_region]
    prices_actual = past_pr1.loc[past_pr1['Year'].astype(str) == y]
    times_use1    = time_pers_final.xs(
        dno_region, level='DNO Region', drop_level=False
    ).xs(reference_year, level='Year', drop_level=False)

    df = apply_bands(
        df, wd_duos_times, 'Weekday', duos_name,
        times_use1, prices_actual
    )
    df = apply_bands(
        df, we_duos_times, 'Weekend', duos_name,
        times_use1, prices_actual
    )

    df.update(df[[duos_name]].ffill())
    df[band_col] = df[band_col].replace('', None)
    df[band_col] = df[band_col].ffill().fillna('')
    df = df.drop(columns=['Time', 'Weekday'])
    df = df.set_index(f'Year {y}')
    _float_cols = [c for c in df.columns if c != band_col]
    df[_float_cols] = df[_float_cols].astype(float)
    df = df.reset_index()

    curr_yr_out = curr_yr_out.merge(
        df, how='outer', left_index=True, right_index=True
    )
    return curr_yr_out

elexon_prices = load_elexon_hh_prices(
    elexon_prices_raw
)  # Cleaned half-hourly Elexon wholesale commodity price table used throughout the Generic profile build

wholesale_cutoff = elexon_prices['DateTime'].max()
print(f"Wholesale data: {elexon_prices['DateTime'].min().date()} to {wholesale_cutoff.date()}")

# =============================================================================
# DUoS TIME BAND CATEGORY TABLES
# =============================================================================

wd_data = pd.Series([
    'Gr1','Gr2','Am1','Am2','Re1','Re2',
    'Am3','Am4','Re3','Re4','Am5','Am6','Gr3','Gr4'
])
we_data = pd.Series([
    'Gr1','Gr2','Am1','Am2','Gr3','Gr4',
    'Am3','Am4','Gr5','Gr6'
])

wd_duos_times = pd.DataFrame(data=wd_data, columns=['Categories'])
we_duos_times = pd.DataFrame(data=we_data, columns=['Categories'])
tarr_cats     = ['Green', 'Amber', 'Red']

def apply_price(x):
    """Maps a band label (e.g. 'Gr1', 'Am3') to its colour category ('Green', 'Amber', 'Red')."""
    for ta in tarr_cats:
        if ta[:2] in x:
            return ta

wd_duos_times['Price Category'] = wd_duos_times['Categories'].apply(
    apply_price
)
we_duos_times['Price Category'] = we_duos_times['Categories'].apply(
    apply_price
)

# =============================================================================
# PRE-COMPUTATION — region-invariant and band-invariant lookups
# All entries built once here; reused across all vol_type × band × region runs.
# =============================================================================

years = pd.Series(range(start_year, build_end_year + 1)).astype(str)
partial_year_info = {}  # maps year string → n_days covered (None = full year)

print("Pre-computing DNO time periods...")
time_pers_final = fetch_time_periods(run_dno_regions)

print("Pre-computing year base frames (date range, tariffs, BSUoS, wholesale, CM)...")
year_base_frames = {}
for _y in years:
    _yi   = int(_y)
    _ynxt = str(_yi + 1)
    _dr   = pd.date_range(start=f"{_y}-04-01", end=f"{_ynxt}-04-01", freq='30min')[:-1]
    _fr   = _dr.to_frame(name=f'Year {_y}')
    _fr['Month']    = _fr[f'Year {_y}'].dt.month
    _fr['Day']      = _fr[f'Year {_y}'].dt.day
    _fr['Year_int'] = _fr[f'Year {_y}'].dt.year
    _fr['Time']     = _fr[f'Year {_y}'].dt.time
    _fr['Weekday']  = _fr[f'Year {_y}'].dt.weekday

    # ── Policy Levy Costs ────────────────────────────────────────────────────
    # Four quarterly rates per component; Q1=Apr–Jun, Q2=Jul–Sep, Q3=Oct–Dec, Q4=Jan–Mar
    _pl_use = pl_tariff_table.xs(_yi, level='year', drop_level=False)
    _pl_comp_cols_y = []
    for _pl_col in pl_tariff_cols:
        _hh_pl_col = f'{_pl_col} [p/kWh] {_y}'
        _pl_comp_cols_y.append(_hh_pl_col)
        _fr[_hh_pl_col] = np.select(
            [_fr['Month'].isin([4,5,6]),    # Q1 Apr–Jun
             _fr['Month'].isin([7,8,9]),    # Q2 Jul–Sep
             _fr['Month'].isin([10,11,12]), # Q3 Oct–Dec
             _fr['Month'].isin([1,2,3])],   # Q4 Jan–Mar
            [_pl_use.xs('q1', level='quarter')[_pl_col].iat[0],
             _pl_use.xs('q2', level='quarter')[_pl_col].iat[0],
             _pl_use.xs('q3', level='quarter')[_pl_col].iat[0],
             _pl_use.xs('q4', level='quarter')[_pl_col].iat[0]]
        )
    _fr[f'Policy Levy Costs [p/kWh] {_y}'] = _fr[_pl_comp_cols_y].sum(axis=1)


    # ── Miscellaneous Charges ─────────────────────────────────────────────────
    # Two seasonal rates per component: summer (Apr–Sep) and winter (Oct–Mar)
    _misc_use = misc_charge_table.xs(_yi, level='year', drop_level=False)
    _misc_comp_cols_y = []
    for _mc in misc_charge_cols:
        _hh_mc_col = f'{_mc} [p/kWh] {_y}'
        _misc_comp_cols_y.append(_hh_mc_col)
        _fr[_hh_mc_col] = np.where(
            _fr['Month'].isin([4,5,6,7,8,9]),  # summer: Apr–Sep
            _misc_use.xs('summer', level='season')[_mc].iat[0],
            _misc_use.xs('winter', level='season')[_mc].iat[0]   # winter: Oct–Mar
        )
    _fr[f'Miscellaneous Charges [p/kWh] {_y}'] = _fr[_misc_comp_cols_y].sum(axis=1)


    # ── BSUoS ─────────────────────────────────────────────────────────────────
    # Single seasonal rate: summer (Apr–Sep) / winter (Oct–Mar); £/MWh ÷ 10 → p/kWh
    _fr[f'BSUoS [p/kWh] {_y}'] = np.where(
        _fr['Month'].isin([4,5,6,7,8,9]),  # summer: Apr–Sep
        bsuos_rates.loc[(_yi, 'summer'), 'bsuos_rate_mwh'] / 10,
        bsuos_rates.loc[(_yi, 'winter'), 'bsuos_rate_mwh'] / 10  # winter: Oct–Mar
    )


    # ── Wholesale Commodity (Elexon HH) ──────────────────────────────────────
    # Clip the frame to the available wholesale horizon for partial years.
    _fy_last_slot = pd.Timestamp(f"{int(_y)+1}-04-01") - pd.Timedelta(minutes=30)
    if wholesale_cutoff < _fy_last_slot:
        _full_len = len(_fr)
        _fr = _fr[_fr[f'Year {_y}'] <= wholesale_cutoff].copy()
        _n_days = int(
            (_fr[f'Year {_y}'].dt.date.max()
             - pd.Timestamp(f"{_y}-04-01").date()).days + 1
        )
        partial_year_info[_y] = _n_days
        print(
            f"  Partial FY {_y}/{int(_y)+1}: "
            f"{_fr[f'Year {_y}'].dt.date.min()} to "
            f"{_fr[f'Year {_y}'].dt.date.max()} "
            f"({len(_fr)} of {_full_len} HH slots, ~{_n_days} days)"
        )
    else:
        partial_year_info[_y] = None  # full year

    _fr = _fr.merge(elexon_prices, how='left', left_on=f'Year {_y}', right_on='DateTime')
    if _fr['Wholesale Commodity [p/kWh]'].isna().any():
        _miss = _fr[_fr['Wholesale Commodity [p/kWh]'].isna()][[f'Year {_y}']].copy()
        raise ValueError(
            f"Gaps in Elexon wholesale prices within available data for FY {_y} "
            f"({len(_miss)} missing rows). Check compiled_elexon_mid_prices.xlsx."
        )
    _fr[f'Wholesale Commodity [p/kWh] {_y}'] = _fr['Wholesale Commodity [p/kWh]']


    # ── Capacity Market ───────────────────────────────────────────────────────
    # Quarterly rate; applied only to weekday evening peak (16:00–19:00), zero otherwise
    _cm_use = cm_table.xs(_yi, level='year', drop_level=False)
    _cm_rates = np.select(
        [_fr['Month'].isin([4,5,6]),    # Q1 Apr–Jun
         _fr['Month'].isin([7,8,9]),    # Q2 Jul–Sep
         _fr['Month'].isin([10,11,12]), # Q3 Oct–Dec
         _fr['Month'].isin([1,2,3])],   # Q4 Jan–Mar
        [_cm_use.xs('q1', level='quarter')['cm'].iat[0],
         _cm_use.xs('q2', level='quarter')['cm'].iat[0],
         _cm_use.xs('q3', level='quarter')['cm'].iat[0],
         _cm_use.xs('q4', level='quarter')['cm'].iat[0]]
    )
    _fr[f'Capacity Market [p/kWh] {_y}'] = np.where(
        (_fr['Weekday'] < 5) & (_fr['Time'] >= time(16, 0)) & (_fr['Time'] < time(19, 0)),
        _cm_rates, 0
    )

    _fr = _fr.drop(columns=['Month', 'Day', 'Year_int', 'DateTime', 'Wholesale Commodity [p/kWh]'])
    year_base_frames[_y] = _fr

print(f"Pre-computation complete: {len(year_base_frames)} year frames built.\n")

# =============================================================================
# MAIN EXECUTION LOOP
# Generic mode only — all voltage types, all bands, all DNO regions
# =============================================================================

for current_vol_type in run_vol_types:

    # ── Voltage-type initialisation ───────────────────────────────────────────
    print(f"\n{'='*60}")
    print(f"Voltage Type: {current_vol_type}")
    print(f"{'='*60}")

    # DUoS volumetric rate chart is voltage-type level — plotted once using
    # bands_to_run[0] as representative (rates vary by band for 2022+ but a
    # single band gives a meaningful cross-regional time-series view).
    _plot_pr = fetch_prices(current_vol_type, run_dno_regions, bands_to_run[0])
    plot_dno_prices(
        _plot_pr, current_vol_type,
        f"{output_root}/{current_vol_type}", run_dno_regions
    )

    for current_band in bands_to_run:

        # ── Band initialisation ───────────────────────────────────────────────
        print(f"\n  Band: {current_band}")

        past_pr    = fetch_prices(current_vol_type, run_dno_regions, current_band)
        dno_prices = past_pr

        for dno_reg in run_dno_regions:

            # ── Region / year processing ──────────────────────────────────────
            dno_region    = dno_reg[:-5]  # strip .xlsx extension
            out_dir       = (
                f"{output_root}/{current_vol_type}/"
                f"{current_band}/1 - Price Profiles/Half-Hourly Prices/"
            )
            parquet_path  = f"{out_dir}{dno_region}.parquet"

            if os.path.exists(parquet_path):
                # Rebuild only the latest year — load the existing parquet, drop
                # any stale latest-year columns, recompute them, and save back.
                latest_y = years.iloc[-1]
                curr_yr_out = pd.read_parquet(parquet_path)
                drop_cols = [
                    c for c in curr_yr_out.columns if c.endswith(f' {latest_y}')
                ]
                if drop_cols:
                    curr_yr_out = curr_yr_out.drop(columns=drop_cols)
                print(f"    {dno_region} — rebuilding FY {latest_y}")
                curr_yr = year_base_frames[latest_y].copy()
                curr_yr_out = voltage_apply(
                    dno_prices, curr_yr, curr_yr_out,
                    current_vol_type, latest_y, dno_region, build_end_year_str,
                    time_pers_final, wd_duos_times, we_duos_times
                )
                curr_yr_out.to_parquet(parquet_path)
                print(f"    ✓ {dno_region} — {current_band} (FY {latest_y} updated)")
                continue

            curr_yr_out = pd.DataFrame(index=range(0, 17570))  # 48 HH slots × 365 days + leap-day buffer
            print(f"    {dno_region}")

            for y in list(years):
                curr_yr = year_base_frames[y].copy()
                curr_yr_out = voltage_apply(
                    dno_prices, curr_yr, curr_yr_out,
                    current_vol_type, y, dno_region, build_end_year_str,
                    time_pers_final, wd_duos_times, we_duos_times
                )

            os.makedirs(out_dir, exist_ok=True)
            curr_yr_out.to_parquet(parquet_path)
            print(f"    ✓ {dno_region} — {current_band}")

# --- END TIMER ---
end_time               = pytime.perf_counter()
execution_time_seconds = end_time - start_time
minutes = int(execution_time_seconds // 60)
seconds = execution_time_seconds % 60
print(
    f"\nTotal programme run time: "
    f"{minutes} minutes and {seconds:.2f} seconds"
)