# -*- coding: utf-8 -*-
"""
Generate TOU Shift Profiles
============================
Generates 12 pre-shifted demand profile workbooks from a baseline half-hourly
demand profile, for use as Manual_Input.py `demand_profile_file` inputs in a
Time-of-Use (TOU) demand-shifting case study run in Post Analysis Specific mode.

Shift mechanic (flat percentage):
  For each half-hour slot T in the Red window (16:00-18:30):
    excess = SHIFT_PCT x original_demand[T]
    demand[T]     -= excess          (source slot is reduced)
    demand[T + N] += excess          (destination slot is increased)
  Excesses are computed from the ORIGINAL (unmodified) demand values for the day
  to avoid cascading effects when a destination slot is also within the Red window.
  Total daily kWh is unchanged by construction.

Scenarios produced (12 files):
  Direction   Hours   10%                                   20%   30%
  Fwd  +1h    +1      <BaselineName> TOU Shift Fwd1h 10pct.xlsx  ...
  Back -1h    -1      <BaselineName> TOU Shift Back1h 10pct.xlsx ...
  Back -2h    -2      <BaselineName> TOU Shift Back2h 10pct.xlsx ...
  Back -3h    -3      <BaselineName> TOU Shift Back3h 10pct.xlsx ...

Note on how much of each Red-band shift exits the Red window:
  Back -3h: all 6 Red slots (16:00-18:30) land outside Red -> full exit
  Back -2h: 4 of 6 slots exit (18:00 and 18:30 stay in Red)
  Back -1h: 2 of 6 slots exit (16:00 and 16:30 only)
  Fwd  +1h: 2 of 6 slots exit (18:00 and 18:30 -> 19:00/19:30)

Configuration:
  Set BASELINE_FILENAME below to the name of your demand profile file
  (must be in Input/Case Study Demand Profiles/).
  Set SCENARIO_PREFIX to a short label used in the output filenames.

Output files are written to:
  Input/Case Study Demand Profiles/

Run this script once before running the TOU Post Analysis scenarios.
Requires: pandas, openpyxl
"""

import os
import sys
import pandas as pd
from datetime import time

# ── User configuration ────────────────────────────────────────────────────────
# Filename of the baseline demand profile (must be in Input/Case Study Demand Profiles/)
BASELINE_FILENAME = 'Imperial College London 2016 Demand Profile.xlsx'

# Short prefix used in output filenames: e.g. 'ICL 2016' -> 'ICL 2016 TOU Shift Fwd1h 10pct.xlsx'
SCENARIO_PREFIX = 'ICL 2016'

# ── Path configuration ────────────────────────────────────────────────────────
# This script is expected to live in the Code Run Folder alongside the main
# pipeline scripts. Paths are resolved relative to this script's location.
CODE_DIR      = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
BASELINE_FILE = os.path.join(CODE_DIR, 'Input', 'Case Study Demand Profiles', BASELINE_FILENAME)
OUTPUT_DIR    = os.path.join(CODE_DIR, 'Input', 'Case Study Demand Profiles')

# ── Constants ─────────────────────────────────────────────────────────────────
HEADER_ROW = 4        # 0-indexed row containing column headers (row 5 in Excel)
RED_START  = time(16, 0)
RED_END    = time(19, 0)   # exclusive: covers 16:00, 16:30, ..., 18:30

SHIFT_DIRECTIONS = {
    'Fwd1h':  +1,
    'Back1h': -1,
    'Back2h': -2,
    'Back3h': -3,
}

SHIFT_MAGNITUDES = [0.10, 0.20, 0.30]   # fractions of Red-band demand to relocate


def load_baseline(path):
    """
    Load a demand profile workbook; return a DataFrame with a DatetimeIndex.
    Date column may be Python datetime objects or Excel serial floats.
    """
    df = pd.read_excel(path, header=HEADER_ROW)
    excel_epoch = pd.Timestamp('1899-12-30')

    def _parse_date(v):
        if pd.isna(v):
            return pd.NaT
        if isinstance(v, (int, float)):
            return excel_epoch + pd.Timedelta(days=float(v))
        return pd.Timestamp(v)

    df['Date'] = df['Date'].apply(_parse_date)
    df = df.dropna(subset=['Date'])
    df = df.set_index('Date').sort_index()
    return df


def shift_one_day_pct(day_df, shift_hours, pct, demand_col='Overall'):
    """
    Apply a flat-percentage Red-band shift to one calendar day.

    For each Red-band slot T: take `pct` of its original demand and add it
    to T + shift_hours. Excesses are computed from original values first so
    that when a destination slot is itself a Red-band slot, we do not
    double-count the source.

    Parameters
    ----------
    day_df      : DataFrame  half-hourly rows for one day, DatetimeIndex
    shift_hours : int        hours to shift (e.g. +1 or -1, -2, -3)
    pct         : float      fraction to relocate (e.g. 0.10 for 10%)
    demand_col  : str        column to modify

    Returns
    -------
    DataFrame   same shape/index, demand_col modified; all other cols unchanged.
    """
    result = day_df.copy()
    result[demand_col] = result[demand_col].astype(float)
    idx = result.index

    red_mask = (idx.time >= RED_START) & (idx.time < RED_END)
    red_idx  = idx[red_mask]

    if red_idx.empty:
        return result

    excess = day_df.loc[red_idx, demand_col].astype(float) * pct

    result.loc[red_idx, demand_col] -= excess

    shift_td = pd.Timedelta(hours=shift_hours)
    dest_idx = red_idx + shift_td

    for dst, exc_kw in zip(dest_idx, excess.values):
        if dst in result.index:
            result.loc[dst, demand_col] += exc_kw
        else:
            # Destination falls outside the day (past midnight or before 00:00);
            # return the excess to the source slot to conserve daily energy.
            src = dst - shift_td
            result.loc[src, demand_col] += exc_kw

    return result


def generate_shifted_profile(baseline_df, shift_hours, pct, demand_col='Overall'):
    """Apply the shift to every calendar day; return a modified copy of baseline_df."""
    shifted = baseline_df.copy()
    shifted[demand_col] = shifted[demand_col].astype(float)
    dates = baseline_df.index.date

    for date, day_df in baseline_df.groupby(dates):
        shifted_day = shift_one_day_pct(day_df, shift_hours, pct, demand_col)
        mask = dates == date
        shifted.loc[mask, demand_col] = shifted_day[demand_col].values

    return shifted


def verify_conservation(original, shifted, demand_col='Overall', tol=1.0):
    """Confirm total kWh is conserved to within tolerance (kW half-hour intervals)."""
    diff = abs(original[demand_col].sum() - shifted[demand_col].sum())
    if diff > tol:
        print(f"  WARNING: energy not conserved — diff = {diff:.3f} kW·intervals")
    else:
        print(f"  Conservation OK (diff = {diff:.4f} kW·intervals)")


def red_band_kwh(df, demand_col='Overall'):
    """Total kWh consumed during 16:00-19:00 (Red band)."""
    idx  = df.index
    mask = (idx.time >= RED_START) & (idx.time < RED_END)
    return df.loc[mask, demand_col].sum() * 0.5


def save_profile(shifted_df, scenario_label, output_dir, header_row, original_path):
    """
    Save a shifted profile in the same proforma format as the baseline.
    Preserves header rows above the column-name row.
    """
    out_path = os.path.join(output_dir, f'{SCENARIO_PREFIX} TOU Shift {scenario_label}.xlsx')

    raw           = pd.read_excel(original_path, header=None)
    header_meta   = raw.iloc[:header_row]
    col_names_row = raw.iloc[[header_row]]
    data_section  = shifted_df.reset_index()

    with pd.ExcelWriter(out_path, engine='openpyxl') as writer:
        header_meta.to_excel(writer, sheet_name='Sheet1',
                             index=False, header=False, startrow=0)
        col_names_row.to_excel(writer, sheet_name='Sheet1',
                               index=False, header=False, startrow=header_row)
        data_section.to_excel(writer, sheet_name='Sheet1',
                              index=False, header=False, startrow=header_row + 1)

    print(f"  Saved: {out_path}")


def main():
    print(f"Loading baseline: {BASELINE_FILE}")
    if not os.path.exists(BASELINE_FILE):
        print(f"ERROR: baseline file not found at {BASELINE_FILE}")
        sys.exit(1)

    baseline_df  = load_baseline(BASELINE_FILE)
    baseline_kwh = baseline_df['Overall'].sum() * 0.5
    baseline_red = red_band_kwh(baseline_df)

    print(f"Baseline: {len(baseline_df)} HH rows | "
          f"Annual kWh = {baseline_kwh:,.0f} | "
          f"Red-band kWh = {baseline_red:,.0f} ({baseline_red/baseline_kwh*100:.1f}% of total)")
    print(f"Output directory: {OUTPUT_DIR}\n")

    os.makedirs(OUTPUT_DIR, exist_ok=True)

    for direction, hours in SHIFT_DIRECTIONS.items():
        for pct in SHIFT_MAGNITUDES:
            pct_label      = f'{int(pct * 100)}pct'
            scenario_label = f'{direction} {pct_label}'
            dir_str        = f'+{hours}h' if hours > 0 else f'{hours}h'

            print(f"Generating {scenario_label} ({dir_str}, {pct*100:.0f}% of Red-band demand)...")
            shifted = generate_shifted_profile(baseline_df, hours, pct, demand_col='Overall')
            verify_conservation(baseline_df, shifted)

            shifted_red   = red_band_kwh(shifted)
            kwh_relocated = baseline_red - shifted_red
            print(f"  Red kWh: {baseline_red:,.0f} -> {shifted_red:,.0f} "
                  f"(relocated {kwh_relocated:,.0f} kWh, "
                  f"{kwh_relocated/baseline_kwh*100:.2f}% of annual)")

            save_profile(shifted, scenario_label, OUTPUT_DIR, HEADER_ROW, BASELINE_FILE)
            print()

    print(f"All {len(SHIFT_DIRECTIONS) * len(SHIFT_MAGNITUDES)} scenarios generated.")


if __name__ == '__main__':
    main()
