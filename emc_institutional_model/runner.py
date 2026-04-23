"""High-level `run_model` API."""

from __future__ import annotations

from dataclasses import dataclass, field

import pandas as pd

from emc_institutional_model.cashflow import (
    build_monthly_cashflows,
    irr_annual_from_monthly_net,
    npv_annual_discount,
    payback_month,
)
from emc_institutional_model.capex import total_capex_usd
from emc_institutional_model.metrics import scenario_kpis
from emc_institutional_model.params import ModelParams
from emc_institutional_model.traditional import (
    traditional_capex_usd,
    traditional_monthly_opex_usd,
    traditional_monthly_revenue_usd,
)


@dataclass
class ModelResult:
    params: ModelParams
    monthly: pd.DataFrame
    kpis: dict[str, float | str | int | None]
    npv_usd: float
    irr_annual: float | str
    payback: int | str
    traditional_summary: dict[str, float] = field(default_factory=dict)


def run_model(p: ModelParams) -> ModelResult:
    rows = build_monthly_cashflows(p)
    df = pd.DataFrame([r.__dict__ for r in rows])
    flows = [r.net_cashflow for r in rows]
    npv_v = npv_annual_discount(flows, p.discount_rate_annual)

    capex = total_capex_usd(p)
    # IRR uses constant month-1 net (matches Scenarios RATE idiom; escalation path is in full monthly series).
    rev0 = rows[0].revenue_inflow if rows else 0.0
    opex0 = rows[0].opex_total if rows else 0.0
    net_flat = rev0 - opex0
    irr_v = irr_annual_from_monthly_net(capex, net_flat, p.analysis_length_months)
    pb = payback_month(rows)
    pb_out: int | str = pb if isinstance(pb, int) else "NO_PAYBACK"

    kpis = scenario_kpis(p)
    kpis["npv_usd"] = npv_v
    kpis["irr_annual_model"] = irr_v

    trad = {
        "traditional_capex_usd": traditional_capex_usd(p),
        "traditional_monthly_opex_usd": traditional_monthly_opex_usd(p),
        "traditional_monthly_revenue_usd": traditional_monthly_revenue_usd(p),
    }

    return ModelResult(
        params=p,
        monthly=df,
        kpis=kpis,
        npv_usd=npv_v,
        irr_annual=irr_v,
        payback=pb_out,
        traditional_summary=trad,
    )
