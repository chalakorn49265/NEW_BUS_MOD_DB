"""kWh algebra (Inputs_Project derived rows in MOZ v2)."""

from __future__ import annotations

from emc_institutional_model.defaults import PRODUCT_ROWS
from emc_institutional_model.params import ModelParams


def product_row(product_type: str) -> tuple[int, int, int, float, float, float, float, float, float, float, int, float, float]:
    return PRODUCT_ROWS[product_type]


def implied_energy_savings_fraction(product_type: str) -> float:
    """(baseline_kWh - delivered_kWh) / baseline_kWh from product row factors."""
    pr = product_row(product_type)
    baseline_f = pr[12]
    delivered_f = pr[11]
    if baseline_f <= 0:
        return 0.0
    return max(0.0, min(1.0 - 1e-9, 1.0 - delivered_f / baseline_f))


def effective_energy_savings_fraction(p: ModelParams) -> float:
    if p.energy_savings_fraction is not None:
        return float(p.energy_savings_fraction)
    return implied_energy_savings_fraction(p.product_type)


def effective_delivered_kwh_factor(p: ModelParams) -> float:
    """Delivered kWh factor after optional savings override (baseline_f * (1 - savings))."""
    pr = product_row(p.product_type)
    baseline_f = pr[12]
    s = effective_energy_savings_fraction(p)
    return baseline_f * (1.0 - s)


def baseline_kwh_month(p: ModelParams) -> float:
    """number_of_poles * 30 * hours * kW * baseline_grid_kwh_factor (Excel B21)."""
    pr = product_row(p.product_type)
    baseline_f = pr[12]
    n_poles = p.number_of_poles
    return n_poles * 30.0 * p.operating_hours_per_night * p.power_kw_per_light * baseline_f


def delivered_kwh_month(p: ModelParams) -> float:
    """number_of_poles * 30 * hours * kW * effective delivered factor."""
    d_eff = effective_delivered_kwh_factor(p)
    return p.number_of_poles * 30.0 * p.operating_hours_per_night * p.power_kw_per_light * d_eff


def avoided_kwh_month(p: ModelParams) -> float:
    return max(0.0, baseline_kwh_month(p) - delivered_kwh_month(p))


def kwh_value_basis_month(p: ModelParams) -> float:
    """Excel B24: avoided_kwh -> avoided else delivered."""
    if p.revenue_basis == "avoided_kwh":
        return avoided_kwh_month(p)
    return delivered_kwh_month(p)
