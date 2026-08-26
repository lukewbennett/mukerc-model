# -*- coding: utf-8 -*-
"""
Case Study Analytics
====================
Single-run analytics for case studies (ASC reduction and TOU demand
shifting). All functions operate on data already available within one
execution of Post Analysis.py; none reads files or makes external calls.

Imported by Post Analysis.py. Never run or edited directly by the user.
"""

import pandas as pd
import numpy as np
from datetime import time


# ─────────────────────────────────────────────────────────────────────────────
# Shared by both case studies
# ─────────────────────────────────────────────────────────────────────────────

def duos_band_breakdown(hh_demand_kw, hh_band, hh_duos_rate_p_per_kwh):
    """
    Consumption and cost share by DUoS band (Green / Amber / Red) for one
    region–year combination.

    Parameters
    ----------
    hh_demand_kw          : array-like  HH demand in kW (Series or ndarray)
    hh_band               : array-like  Band label ('Green', 'Amber', 'Red'),
                                        positionally aligned to hh_demand_kw
    hh_duos_rate_p_per_kwh: array-like  DUoS volumetric rate in p/kWh,
                                        positionally aligned to hh_demand_kw

    Returns
    -------
    DataFrame  index = ['Green', 'Amber', 'Red'], columns:
        Consumption [kWh], Consumption Share [%],
        DUoS Volumetric Cost [£], DUoS Cost Share [%]
    """
    # Convert to numpy arrays so arithmetic is always element-wise by position.
    # pandas Series with duplicate DatetimeIndex labels (UK DST fall-back day)
    # produce cartesian products when multiplied together — numpy avoids this.
    demand_a = np.asarray(hh_demand_kw, dtype=float)
    rate_a   = np.asarray(hh_duos_rate_p_per_kwh, dtype=float)
    band_a   = np.asarray(hh_band)

    hh_kwh  = demand_a * 0.5
    hh_cost = hh_kwh * rate_a / 100

    total_kwh  = hh_kwh.sum()
    total_cost = hh_cost.sum()

    rows = {}
    for band in ['Green', 'Amber', 'Red']:
        mask       = band_a == band
        band_kwh   = hh_kwh[mask].sum()
        band_cost  = hh_cost[mask].sum()
        rows[band] = {
            'Consumption [kWh]':        round(band_kwh,  0),
            'Consumption Share [%]':    round(band_kwh  / total_kwh  * 100, 2) if total_kwh  > 0 else 0.0,
            'DUoS Volumetric Cost [£]': round(band_cost, 2),
            'DUoS Cost Share [%]':      round(band_cost / total_cost * 100, 2) if total_cost > 0 else 0.0,
        }

    return pd.DataFrame(rows).T


def blended_duos_rate(hh_demand_kw, hh_duos_rate_p_per_kwh):
    """
    Total DUoS Volumetric Cost [£] / total annual kWh = effective DUoS
    unit rate in p/kWh. Comparable across scenarios and bands.

    Returns
    -------
    float  effective DUoS unit rate in p/kWh (rounded to 4 dp)
    """
    demand_a = np.asarray(hh_demand_kw, dtype=float)
    rate_a   = np.asarray(hh_duos_rate_p_per_kwh, dtype=float)
    hh_kwh     = demand_a * 0.5
    total_kwh  = hh_kwh.sum()
    total_cost = (hh_kwh * rate_a / 100).sum()  # £
    if total_kwh > 0:
        return round(total_cost / total_kwh * 100, 4)  # p/kWh
    return 0.0


def triad_exposure_share(hh_demand_kw, hh_band):
    """
    Of the Red-band half-hourly consumption, the share that also falls within
    the TRIAD window (weekday, 16:00–19:00, November–February).

    This contextualises TOU shift scenarios: if a site shifts demand out of
    Red-band hours, does that shift also reduce TRIAD/CM window exposure?

    Parameters
    ----------
    hh_demand_kw : Series       HH demand in kW, datetime-indexed
    hh_band      : array-like   Band label, positionally aligned to hh_demand_kw

    Returns
    -------
    float  % of Red-band kWh that falls within the TRIAD window (0–100)
    """
    idx = hh_demand_kw.index
    # DatetimeIndex attribute access (.weekday, .time, .month) returns numpy arrays,
    # so triad_mask is a numpy boolean array — safe with duplicate index labels.
    triad_mask = (
        (idx.weekday < 5) &
        (idx.time >= time(16, 0)) &
        (idx.time < time(19, 0)) &
        idx.month.isin([11, 12, 1, 2])
    )
    demand_a = np.asarray(hh_demand_kw, dtype=float)
    band_a   = np.asarray(hh_band)
    red_mask      = band_a == 'Red'
    red_kwh       = (demand_a[red_mask] * 0.5).sum()
    red_triad_kwh = (demand_a[red_mask & triad_mask] * 0.5).sum()
    if red_kwh > 0:
        return round(red_triad_kwh / red_kwh * 100, 2)
    return 0.0


# ─────────────────────────────────────────────────────────────────────────────
# ASC reduction case study
# ─────────────────────────────────────────────────────────────────────────────

def exceeded_cap_incidence(monthly_max_kva, capacity):
    """
    Number and proportion of billing periods (calendar months) in which the
    monthly peak half-hourly demand exceeds the contracted ASC (capacity).

    Parameters
    ----------
    monthly_max_kva : Series  monthly peak kVA, MS-resampled (from _monthly_max)
    capacity        : float   ASC in kVA

    Returns
    -------
    dict  {'months_breached': int, 'months_total': int, 'proportion_breached': float}
    """
    breached    = (monthly_max_kva > capacity).sum()
    total       = len(monthly_max_kva)
    proportion  = round(breached / total * 100, 2) if total > 0 else 0.0
    return {
        'months_breached':     int(breached),
        'months_total':        int(total),
        'proportion_breached': proportion,
    }


def headroom_utilisation(hh_demand_kw, capacity, power_factor):
    """
    Average and peak half-hourly demand as a percentage of the contracted
    capacity (kVA). Quantifies over-provisioning (Scenario A) and remaining
    headroom after ASC reduction (Scenario B).

    Parameters
    ----------
    hh_demand_kw : Series  HH demand in kW, datetime-indexed
    capacity     : float   ASC in kVA
    power_factor : float   site power factor (used to convert kW → kVA)

    Returns
    -------
    dict  {'avg_utilisation_pct': float, 'peak_utilisation_pct': float}
    """
    hh_kva = hh_demand_kw / power_factor
    avg_pct  = round(hh_kva.mean()  / capacity * 100, 2) if capacity > 0 else 0.0
    peak_pct = round(hh_kva.max()   / capacity * 100, 2) if capacity > 0 else 0.0
    return {
        'avg_utilisation_pct':  avg_pct,
        'peak_utilisation_pct': peak_pct,
    }
