#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Helper Functions
================
Script 2 of 4 in the MUKERC pipeline.

Shared utility functions used by Build Price Profiles.py and
Post Analysis.py. Centralising these functions ensures that DNO
tariff row-label logic and annex sheet-name logic remain consistent
across both scripts and only need to be updated in one place when
naming conventions change.

Created on Sun Jun 07 2026
@author: Luke Bennett
"""

import pandas as pd


def get_vol_row_label(vol_type, year, residual_band):
    """
    Returns the correct DNO spreadsheet row label for a given voltage
    type and year combination.

    This function handles Import tariffs only. Export functionality
    has been removed as it is outside the scope of the MUKERC framework.

    Handles four distinct naming convention periods:
      Pre-2021:  Legacy naming ('HH Metered', 'Non-CT' etc.)
      2021:      Transition naming ('Site Specific',
                 'Non-Domestic Aggregated')
      2022-2023: Banded naming (Band 1-4 appended to row label)
      2024+:     Updated banded naming for Non-Domestic Aggregated
                 ('Non-Domestic Aggregated or CT Band X')

    Parameters
    ----------
    vol_type     : string — 'LV Sub', 'LV', 'HV' or
                            'Non-Domestic Aggregated'
    year         : string or int — year e.g. '2022' or 2022
    residual_band: string — e.g. 'Band 4'
    """
    year = int(year)

    if vol_type == 'Non-Domestic Aggregated':
        if year >= 2024:
            # 2024+: label updated to include 'or CT'
            return 'Non-Domestic Aggregated or CT ' + residual_band
        elif year >= 2022:
            # 2022-2023: banded label without 'or CT'
            return 'Non-Domestic Aggregated ' + residual_band
        elif year == 2021:
            return 'Non-Domestic Aggregated'
        else:
            return 'LV Network Non-Domestic Non-CT'

    elif vol_type == 'LV':
        if year >= 2022:
            return 'LV Site Specific ' + residual_band
        elif year == 2021:
            return 'LV Site Specific'
        else:
            return 'LV HH Metered'

    elif vol_type == 'LV Sub':
        if year >= 2022:
            return 'LV Sub Site Specific ' + residual_band
        elif year == 2021:
            return 'LV Sub Site Specific'
        else:
            return 'LV Sub HH Metered'

    elif vol_type == 'HV':
        if year >= 2022:
            return 'HV Site Specific ' + residual_band
        elif year == 2021:
            return 'HV Site Specific'
        else:
            return 'HV HH Metered'


def get_annex_sheet(year):
    """
    Returns the correct Annex 1 sheet name for a given year.

    Parameters
    ----------
    year : string or int — year e.g. '2022' or 2022
    """
    if int(year) >= 2021:
        return 'Annex 1 LV, HV and UMS charges'
    else:
        return 'Annex 1 LV and HV charges'

def get_pl_tariff_component_cols(tariff_df):
    """
    Returns the list of Policy Levy tariff component columns from the
    tariff workbook.

    Assumes column names are already lowercase-normalised. Excludes the
    index/dimension columns (year, season, quarter) and the capacity
    market column (cm), which is extracted separately for peak-window use.
    """
    id_cols = {'year', 'season', 'quarter', 'cm'}
    return [
        col for col in tariff_df.columns
        if col not in id_cols
    ]


def prepare_tariff_tables(tariff_df):
    """
    Prepares Policy Levy tariff input tables for Build Profiles and Post Analysis.

    Expected workbook format (Policy Levy Rates.xlsx):
        Year | Season | Quarter | <Policy Levy components...> | CM

    Quarter values: Q1=Apr-Jun, Q2=Jul-Sep, Q3=Oct-Dec, Q4=Jan-Mar.
    All values in £/MWh — converted to p/kWh internally (÷10).

    Returns:
    - pl_tariff_cols: list of Policy Levy component column names (lowercase)
    - pl_tariff_table: MultiIndex (year, quarter) with each component in p/kWh
    - pl_tariff_total: same index, plus 'Tariff Sum [p/kWh]' column
    - cm_table: MultiIndex (year, quarter) with 'cm' column in p/kWh
    """
    df = tariff_df.copy()
    df.columns = [str(c).strip().lower() for c in df.columns]

    pl_tariff_cols = get_pl_tariff_component_cols(df)

    df['year']    = df['year'].astype(int)
    df['quarter'] = df['quarter'].astype(str).str.strip().str.lower()

    tariff_indexed = df.set_index(['year', 'quarter'])

    pl_tariff_table = tariff_indexed[pl_tariff_cols].copy()
    pl_tariff_table = pl_tariff_table / 10  # £/MWh → p/kWh

    pl_tariff_total = pl_tariff_table.copy()
    pl_tariff_total['Tariff Sum [p/kWh]'] = pl_tariff_total.sum(axis=1)

    cm_table = tariff_indexed[['cm']].copy()
    cm_table = cm_table / 10  # £/MWh → p/kWh

    return pl_tariff_cols, pl_tariff_table, pl_tariff_total, cm_table

def prepare_seasonal_component_tables(component_df):
    """
    Prepares a generic seasonal component workbook for downstream use.

    Assumes the workbook contains:
    - year
    - season
    - all remaining columns as component charges (values in £/MWh)

    All values converted from £/MWh to p/kWh internally (÷10).

    Returns:
    - component_cols: list of component column names (lowercase)
    - component_table: MultiIndex (year, season) with each component
      in p/kWh
    - component_total: same index, plus 'Component Sum [p/kWh]' column
    """
    df = component_df.copy()
    df.columns = [str(c).strip().lower() for c in df.columns]

    id_cols = {'year', 'season'}
    component_cols = [
        col for col in df.columns
        if col not in id_cols
    ]

    df['year']   = df['year'].astype(int)
    df['season'] = df['season'].astype(str).str.strip().str.lower()

    component_indexed = df.set_index(['year', 'season'])

    component_table = component_indexed[component_cols].copy()
    component_table = component_table / 10  # £/MWh → p/kWh

    component_total = component_table.copy()
    component_total['Component Sum [p/kWh]'] = component_total.sum(axis=1)

    return component_cols, component_table, component_total


def load_bsuos_rates(path):
    """
    Loads seasonal BSUoS rates from the Balancing input workbook.

    Expected format (header at row 3, same convention as TNUoS sheets):
        Year | Season | BSUoS Tariff (£/MWh)

    Values are in £/MWh. Divide by 10 at the call site to convert to p/kWh.

    Returns a DataFrame indexed on (year, season) with one column
    'bsuos_rate_mwh'.
    """
    df = pd.read_excel(path, header=3)
    df.columns = [str(c).strip() for c in df.columns]
    df.columns = ['year', 'season', 'bsuos_rate_mwh']
    df = df.dropna(subset=['year'])
    df['year']   = df['year'].astype(int)
    df['season'] = df['season'].str.strip().str.lower()
    return df.set_index(['year', 'season'])
