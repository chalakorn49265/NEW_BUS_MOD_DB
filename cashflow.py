"""Monthly cashflow, payback, IRR, NPV (Cashflow + ROI_Payback + Scenarios IRR)."""

from __future__ import annotations

from dataclasses import dataclass
from typing import Literal, Union

import numpy_financial as npf

from emc_institutional_model.capex import total_capex_usd
from emc_institutional_model.energy import kwh_value_basis_month
from emc_institutional_model.opex import opex_month1
from emc_institutional_model.params import ModelParams
from emc_institutional_model.traditional import traditional_electrical_component_monthly


@dataclass(frozen=True)
class MonthRow:
    month_index: int
    capex_outflow: float
    electrical_fee: float
    maintenance_fee: float
    opex_total: float
    revenue_energy: float
    revenue_custody: float
    revenue_performance_fee: float
    revenue_inflow: float
    tax_cash: float
    distribution_fees: float
    net_cashflow: float
    cumulative_net_cashflow: float
    net_cashflow_adjusted: float
    cumulative_net_adjusted: float


def _esc(p: ModelParams, month_index: int) -> float:
    return (1.0 + p.escalation_pct_annual / 12.0) ** (month_index - 1)


def build_monthly_cashflows(p: ModelParams) -> list[MonthRow]:
    """Replicates Cashflow sheet rows for months 1..analysis_length_months.

    ``net_cashflow`` / ``cumulative_net_cashflow`` are the **equipment project** series (full CAPEX outflow month 1).
    ``net_cashflow_adjusted`` adds optional cash tax and distributions for Sources & Uses views;
    with default ``EmcAdjustments`` it equals ``net_cashflow``.
    """
    capex = total_capex_usd(p)
    o1 = opex_month1(p)
    kwh_basis = kwh_value_basis_month(p)
    trad_elec_m1 = traditional_electrical_component_monthly(p)
    adj = p.emc_adjustments
    dep_m = capex / max(1, int(adj.depreciation_months))
    tax_rate = adj.corporate_tax_rate

    rows: list[MonthRow] = []
    cum = 0.0
    cum_adj = 0.0
    for m in range(1, p.analysis_length_months + 1):
        esc = _esc(p, m)
        capex_o = -capex if m == 1 else 0.0
        elec = -o1.electrical_fee * esc
        maint = -o1.maintenance_fee * esc
        opex_tot = elec + maint
        rev_e = p.tariff_model.gov_payment.energy_usd(kwh_basis) * esc
        rev_c = (
            (p.custody_fee_usd_per_pole_month * p.number_of_poles * esc)
            if p.custody_fee_enabled
            else 0.0
        )
        perf_base = p.emc_performance_fee_pct_of_energy_savings * max(
            0.0, trad_elec_m1 - o1.electrical_fee
        )
        rev_p = perf_base * esc
        rev = rev_e + rev_c + rev_p
        net = capex_o + opex_tot + rev
        cum += net

        dep = dep_m if m <= int(adj.depreciation_months) else 0.0
        taxable = rev_e + rev_p + elec + maint - dep
        tax_c = -tax_rate * max(0.0, taxable)

        gross_in = rev_e + rev_c + rev_p
        dist = -adj.distribution_pct_of_gross_inflow * gross_in - adj.distribution_fixed_usd_month

        net_adj = net + tax_c + dist
        cum_adj += net_adj

        rows.append(
            MonthRow(
                month_index=m,
                capex_outflow=capex_o,
                electrical_fee=elec,
                maintenance_fee=maint,
                opex_total=opex_tot,
                revenue_energy=rev_e,
                revenue_custody=rev_c,
                revenue_performance_fee=rev_p,
                revenue_inflow=rev,
                tax_cash=tax_c,
                distribution_fees=dist,
                net_cashflow=net,
                cumulative_net_cashflow=cum,
                net_cashflow_adjusted=net_adj,
                cumulative_net_adjusted=cum_adj,
            )
        )
    return rows


def payback_month(rows: list[MonthRow]) -> Union[int, Literal["NO_PAYBACK"]]:
    for r in rows:
        if r.cumulative_net_cashflow >= 0:
            return r.month_index
    return "NO_PAYBACK"


def payback_month_adjusted(rows: list[MonthRow]) -> Union[int, Literal["NO_PAYBACK"]]:
    for r in rows:
        if r.cumulative_net_adjusted >= 0:
            return r.month_index
    return "NO_PAYBACK"


def irr_annual_from_monthly_net(
    capex: float, monthly_net: float, n_months: int
) -> Union[float, Literal["NO_IRR"]]:
    """Excel Scenarios column L: (1+RATE(n, pmt, -pv, 0))^12-1."""
    if monthly_net <= 0 or n_months <= 0:
        return "NO_IRR"
    try:
        r_m = npf.rate(n_months, monthly_net, -capex, 0.0)
    except Exception:
        return "NO_IRR"
    if r_m is None or (isinstance(r_m, float) and r_m != r_m):  # NaN
        return "NO_IRR"
    return float((1.0 + r_m) ** 12 - 1.0)


def npv_annual_discount(monthly_flows: list[float], annual_discount: float) -> float:
    """NPV with end-of-month cashflows; first flow is month 1 (after t=0)."""
    r_m = (1.0 + annual_discount) ** (1.0 / 12.0) - 1.0
    pv = 0.0
    for t, c in enumerate(monthly_flows, start=1):
        pv += c / (1.0 + r_m) ** t
    return pv
