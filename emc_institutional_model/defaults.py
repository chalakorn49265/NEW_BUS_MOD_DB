"""
Default constants copied from the MOZ v2 workbook generator
(`DELIVERABLES/TRACK_C/build_moz_institutional_model.py`); not loaded from .xlsx.
"""

from __future__ import annotations

# Scenario rows: Base, Conservative, Aggressive — labor, transport, trench, security, permitting, install_index
SCENARIO_FACTORS: dict[str, tuple[float, float, float, float, float, float]] = {
    "Base": (1.00, 1.00, 1.00, 1.00, 1.00, 1.00),
    "Conservative": (1.10, 1.10, 1.15, 1.08, 1.12, 1.10),
    "Aggressive": (0.92, 0.95, 0.92, 0.95, 0.94, 0.94),
}

TIER_DEFAULTS: dict[str, tuple[float, float, float, float, float]] = {
    "city_center": (1.18, 1.05, 1.30, 1.10, 30.0),
    "suburb": (1.00, 1.00, 1.00, 1.00, 20.0),
    "rural": (0.90, 1.25, 1.45, 1.20, 25.0),
}

# product_type -> (is_grid, req_ug, has_batt, fixture, trench, cable, labor, logistics,
#                  routine_maint, batt_cost, batt_cycle, delivered_f, baseline_f)
PRODUCT_ROWS: dict[str, tuple[int, int, int, float, float, float, float, float, float, float, int, float, float]] = {
    "AI_lightning_grid": (1, 1, 0, 200, 60, 50, 45, 30, 1.20, 0, 0, 0.95, 2.30),
    "AI_battery_integrated_grid": (1, 1, 1, 230, 60, 50, 50, 32, 1.30, 40, 72, 0.85, 2.40),
    "AI_plus_solar_offgrid": (0, 0, 1, 280, 0, 0, 50, 35, 1.00, 70, 72, 1.00, 2.60),
}

# Traditional_Benchmark sheet B4:B14 order
TRADITIONAL_BENCHMARK: dict[str, float] = {
    "legacy_hw_usd_per_light": 205.0,
    "legacy_trench_usd_per_light": 75.0,
    "legacy_cable_usd_per_light": 58.0,
    "legacy_labor_usd_per_pole": 55.0,
    "legacy_logistics_usd_per_pole": 36.0,
    "trad_edm_kwh_multiplier": 1.20,
    "trad_routine_om_usd_per_pole_month": 2.0,
    "trad_security_coupling_pct": 0.10,
    "trad_baseline_kwh_factor_high": 2.60,
    "trad_delivered_kwh_factor": 1.20,
    # Excel row key: trad_service_fee_revenue_scale (托管费计入传统收入折算)
    "trad_service_fee_revenue_scale": 0.9,
}

DEFAULT_GOV_FLAT_USD_PER_KWH = 0.18
DEFAULT_EDM_FLAT_USD_PER_KWH = 0.10
