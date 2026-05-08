"""KPI helpers aligned with Scenarios sheet."""

from __future__ import annotations

from emc_institutional_model.capex import total_capex_usd
from emc_institutional_model.cashflow import build_monthly_cashflows, irr_annual_from_monthly_net
from emc_institutional_model.opex import opex_month1
from emc_institutional_model.energy import kwh_value_basis_month
from emc_institutional_model.params import ModelParams
from emc_institutional_model.traditional import (
    traditional_capex_usd,
    traditional_electrical_component_monthly,
    traditional_monthly_opex_usd,
    traditional_monthly_revenue_usd,
)


def scenario_kpis(p: ModelParams) -> dict[str, float | str | int]:
    """Single-row Scenarios-style metrics for active product and location tier."""
    capex = total_capex_usd(p)
    o1 = opex_month1(p)
    opex_m = o1.total_opex
    kwh_basis = kwh_value_basis_month(p)
    rev_energy = p.tariff_model.gov_payment.energy_usd(kwh_basis)
    rev_custody = (
        p.custody_fee_usd_per_pole_month * p.number_of_poles if p.custody_fee_enabled else 0.0
    )
    perf_m1 = p.emc_performance_fee_pct_of_energy_savings * max(
        0.0, traditional_electrical_component_monthly(p) - o1.electrical_fee
    )
    rev_m = rev_energy + rev_custody + perf_m1
    monthly_net = rev_m - opex_m
    irr = irr_annual_from_monthly_net(capex, monthly_net, p.analysis_length_months)
    rows = build_monthly_cashflows(p)
    from emc_institutional_model.cashflow import payback_month

    pb = payback_month(rows)
    fee_part = rev_custody if p.custody_fee_enabled else 0.0
    req_fee = max(
        0.0,
        (
            capex / max(1, p.analysis_length_months)
            + opex_m
            - (rev_m - fee_part)
        )
        / max(1, p.number_of_poles),
    )

    tc = traditional_capex_usd(p)
    to = traditional_monthly_opex_usd(p)
    tr = traditional_monthly_revenue_usd(p)
    tnet = tr - to
    t_irr = irr_annual_from_monthly_net(tc, tnet, p.analysis_length_months)

    gap = None
    if isinstance(irr, float) and isinstance(t_irr, float):
        gap = irr - t_irr

    return {
        "est_capex_usd": capex,
        "est_monthly_opex_usd": opex_m,
        "est_monthly_revenue_usd": rev_m,
        "est_payback_months": pb,
        "required_custody_fee_usd_per_pole_month": req_fee,
        "est_irr_annual": irr,
        "traditional_capex_usd": tc,
        "traditional_monthly_opex_usd": to,
        "traditional_monthly_revenue_usd": tr,
        "traditional_est_irr_annual": t_irr,
        "irr_gap_ai_minus_traditional": gap,
    }
