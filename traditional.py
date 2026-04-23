"""Traditional (incumbent) lighting benchmark — Scenarios columns M–O logic."""

from __future__ import annotations

from emc_institutional_model.energy import delivered_kwh_month
from emc_institutional_model.location import effective_multipliers
from emc_institutional_model.params import ModelParams


def traditional_capex_usd(p: ModelParams) -> float:
    tb = p.traditional
    labor_m, trans_m, trench_m, _sec = effective_multipliers(p)
    lights = p.number_of_lights
    poles = p.number_of_poles
    return (
        lights * tb["legacy_hw_usd_per_light"]
        + lights * tb["legacy_trench_usd_per_light"] * trench_m
        + lights * tb["legacy_cable_usd_per_light"] * trench_m
        + poles * tb["legacy_labor_usd_per_pole"] * labor_m
        + poles * tb["legacy_logistics_usd_per_pole"] * trans_m
    )


def traditional_monthly_opex_usd(p: ModelParams) -> float:
    tb = p.traditional
    _labor_m, _trans_m, _trench_m, sec_m = effective_multipliers(p)
    edm_kwh = (
        p.operating_hours_per_night
        * 30.0
        * p.power_kw_per_light
        * p.number_of_lights
        * tb["trad_edm_kwh_multiplier"]
    )
    electrical = p.tariff_model.edm_cost.energy_usd(edm_kwh)
    routine = (
        p.number_of_poles
        * tb["trad_routine_om_usd_per_pole_month"]
        * (1.0 + tb["trad_security_coupling_pct"] * (sec_m - 1.0))
    )
    return electrical + routine


def traditional_monthly_revenue_usd(p: ModelParams) -> float:
    """Gov payment on billed kWh proxy + optional custody fee scaled for traditional."""
    tb = p.traditional
    if p.revenue_basis == "avoided_kwh":
        factor = max(0.0, tb["trad_baseline_kwh_factor_high"] - tb["trad_delivered_kwh_factor"])
    else:
        factor = tb["trad_delivered_kwh_factor"]
    kwh_block = (
        p.operating_hours_per_night * 30.0 * p.power_kw_per_light * p.number_of_lights * factor
    )
    energy_rev = p.tariff_model.gov_payment.energy_usd(kwh_block)
    custody = 0.0
    if p.custody_fee_enabled:
        custody = (
            p.custody_fee_usd_per_pole_month
            * p.number_of_poles
            * tb["trad_service_fee_revenue_scale"]
        )
    return energy_rev + custody


def traditional_electrical_component_monthly(p: ModelParams) -> float:
    """Grid EDM cost only (for EMC performance fee on energy savings)."""
    tb = p.traditional
    edm_kwh = (
        p.operating_hours_per_night
        * 30.0
        * p.power_kw_per_light
        * p.number_of_lights
        * tb["trad_edm_kwh_multiplier"]
    )
    return p.tariff_model.edm_cost.energy_usd(edm_kwh)
