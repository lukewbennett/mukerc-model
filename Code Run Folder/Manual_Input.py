#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Manual Input Configuration
==========================
Script 1 of 4 in the MUKERC pipeline.

All user-configurable settings for the MUKERC pipeline are defined here.
Edit this file before each run. Sections are ordered by how often they
need changing: Section 1 (every run) through Section 7 (rarely, if ever).

Created on Sun Jun 07 2026
@author: Luke Bennett
"""

import math
from datetime import time


# =============================================================================
# SECTION 1 — RUN MODE
# Set this first — it controls which other settings are active.
# =============================================================================

# 'Generic'  → Full sweep across all DNO regions, voltage types, and bands.
#              Uses representative flat demand profiles per band (Section 4).
#              Outputs → Output Data Repository/Generic Results/
#
# 'Specific' → Named case study analysis for a real or representative site.
#              Reads price profiles built by Generic mode.
#              Outputs → Output Data Repository/Case Studies/{case_study_name}/
run_mode = 'Generic'


# =============================================================================
# SECTION 2 — YEAR RANGE
# Applies to both Generic and Specific mode.
# start_year must have distribution price data in Input/Network Prices/Distribution/
# =============================================================================

start_year     = 2017   # First financial year (April start)
end_year_limit = 2026   # Last financial year


# =============================================================================
# SECTION 3 — SPECIFIC MODE SETTINGS
# Only edit when run_mode = 'Specific'. Ignored in Generic mode.
# =============================================================================

# ── Case study identity ──────────────────────────────────────────────────────

# Name for this case study — creates a subfolder under Case Studies/
case_study_name = 'ICL 2016 - ASC Below Peak'

# DNO region(s) to analyse — filenames must match those in:
# Input/Network Prices/Distribution/{year}/
specific_regions = [
    '12 - London.xlsx',
]

# Demand profile to use:
#   'Flat'                    → constant flat demand at 'basis' kW all year
#   '<filename>.xlsx'         → Excel workbook in Input/Case Study Demand Profiles/
demand_profile_file = 'Imperial College London 2016 Demand Profile.xlsx'

# ── Voltage connection ───────────────────────────────────────────────────────

# Voltage type and MIC flag — must be a valid combination from the table:
#   ┌─────────────────────────────┬─────────┬──────────────────────────────┐
#   │ vol_type                    │ has_mic │ Description                  │
#   ├─────────────────────────────┼─────────┼──────────────────────────────┤
#   │ 'Non-Domestic Aggregated'   │ False   │ LV connected, no MIC         │
#   │ 'LV'                        │ True    │ LV connected, with MIC/ASC   │
#   │ 'LV Sub'                    │ True    │ LV Sub connected, with MIC   │
#   │ 'HV'                        │ True    │ HV connected, with MIC/ASC   │
#   └─────────────────────────────┴─────────┴──────────────────────────────┘
vol_type = 'HV'
has_mic  = True

# Residual charging band — options: 'Band 1', 'Band 2', 'Band 3', 'Band 4'
residual_band = 'Band 4'

# ── Site electrical parameters ───────────────────────────────────────────────

# Flat demand level — used as the constant load when demand_profile_file = 'Flat'
# and as the TRIAD demand basis for TNUoS estimation in all Specific mode runs.
basis = 11000  # kW

# Assumed power factor — used to derive capacity (kVA) and reactive power (kVAr).
# Applies to Specific mode here. In Generic mode, Post Analysis derives these
# automatically per band from generic_flat_loads and the same power_factor.
power_factor = 0.98

capacity       = basis / power_factor                        # kVA — DUoS capacity charge basis
reactive_power = basis * math.tan(math.acos(power_factor))  # kVAr — DUoS reactive charge basis

# ── Chart settings ───────────────────────────────────────────────────────────

# Peak window shown in price profile bar charts
plot_pk_start = time(16, 0)
plot_pk_end   = time(18, 30)


# =============================================================================
# SECTION 4 — GENERIC MODE REPRESENTATIVE FLAT LOADS
# Flat demand levels [kW] representative of each voltage type and residual
# charging band combination. Derived from the midpoint of annual consumption
# ranges per band; Band 4 uses an engineering upper ceiling.
#
# In Generic mode, Post Analysis derives DUoS capacity and reactive charges
# automatically from these values using power_factor (Section 3).
# =============================================================================

generic_flat_loads = {
    'Non-Domestic Aggregated': {
        'Band 1':   0.20,
        'Band 2':   0.92,
        'Band 3':   2.16,
        'Band 4':   3.61
    },
    'LV Sub': {
        'Band 1':  12.00,
        'Band 2':  34.50,
        'Band 3':  57.15,
        'Band 4':  86.63
    },
    'LV': {
        'Band 1':  12.00,
        'Band 2':  34.50,
        'Band 3':  57.15,
        'Band 4':  86.63
    },
    'HV': {
        'Band 1':   63.30,
        'Band 2':  213.30,
        'Band 3':  420.00,
        'Band 4':  675.00
    }
}


# =============================================================================
# SECTION 5 — BUILD PRICE PROFILES SCOPE CONTROL
# Controls whether Build Price Profiles runs the full Generic scope or a
# reduced test scope for faster iteration and debugging.
# =============================================================================

# 'Full' → all voltage types, all DNO regions, all four bands
# 'Test' → restricted to the three lists below
build_scope_mode = 'Full'

build_test_vol_types   = ['HV']
build_test_dno_regions = ['12 - London.xlsx']
build_test_bands       = ['Band 4']


# =============================================================================
# SECTION 6 — OUTPUT OPTIONS
# Controls optional output layers added on top of the core ex-VAT decomposition.
# =============================================================================

# When True, a 'Total Cost incl. VAT [£]' column (standard rate: 20%) is added
# to detailed price breakdown outputs alongside the existing ex-VAT total.
# The ex-VAT decomposition is always preserved — VAT is an additional layer.
# Default: False — VAT is recoverable by VAT-registered non-domestic consumers;
# ex-VAT figures are standard for structural cost comparisons.
apply_vat = False


# =============================================================================
# SECTION 7 — FRAMEWORK REFERENCE LISTS
# Fixed lists used internally by both scripts. Only edit if DNO region
# filenames or voltage type names change in the underlying input data.
# =============================================================================

all_dno_regions = [
    '10 - Eastern.xlsx',
    '11 - East Midlands.xlsx',
    '12 - London.xlsx',
    '13 - N Wales & Mersey.xlsx',
    '14 - Midlands.xlsx',
    '15 - North East.xlsx',
    '16 - North West.xlsx',
    '17 - Northern Scotland.xlsx',
    '18 - Southern Scotland.xlsx',
    '19 - South East.xlsx',
    '20 - Southern.xlsx',
    '21 - South Wales.xlsx',
    '22 - South Western.xlsx',
    '23 - Yorkshire.xlsx'
]

all_vol_types = [
    'Non-Domestic Aggregated',
    'LV Sub',
    'LV',
    'HV'
]

