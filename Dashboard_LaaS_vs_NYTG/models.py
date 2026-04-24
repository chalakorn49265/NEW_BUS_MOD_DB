from __future__ import annotations

from dataclasses import dataclass
from typing import Literal

from business_model_comparison.metrics import (
    DebtMetrics,
    compute_dscr_by_year,
    irr_annual_from_monthly_cashflows,
    npv_monthly,
    payback_month_from_monthly_cashflows,
)
from business_model_comparison.provenance import Provenance, SeriesWithProv, SourceRef
from business_model_comparison.roadlight_data import RoadlightParsed, to_rmb


@dataclass(frozen=True)
class BaselineResults:
    years: list[int]
    # Annual P&L style (RMB/year)
    revenue_rmb_y: SeriesWithProv
    cash_opex_rmb_y: SeriesWithProv
    electricity_opex_rmb_y: SeriesWithProv
    depreciation_rmb_y: SeriesWithProv
    accounting_gross_profit_rmb_y: SeriesWithProv
    # Debt
    debt_principal_rmb_y: SeriesWithProv
    debt_interest_rmb_y: SeriesWithProv
    debt_service_rmb_y: SeriesWithProv
    dscr: DebtMetrics
    # Cashflows monthly (month0 index)
    project_cashflows_month0: list[float]
    equity_cashflows_month0: list[float]
    # KPIs
    payback_months: int | Literal["NO_PAYBACK"]
    irr_project_annual: float | Literal["NO_IRR"]
    irr_equity_annual: float | Literal["NO_IRR"]
    npv_project_rmb: float
    npv_equity_rmb: float
    capex_y0_rmb: float
    capex_provenance: Provenance


def _series_sub(a: SeriesWithProv, b: SeriesWithProv, *, units: str, transform: str, sources: list[SourceRef]) -> SeriesWithProv:
    years = sorted(set(a.values_by_year.keys()) | set(b.values_by_year.keys()))
    return SeriesWithProv(
        values_by_year={y: float(a.get(y)) - float(b.get(y)) for y in years},
        provenance=Provenance(tuple(sources), units=units, transform=transform),
    )


def _series_add(a: SeriesWithProv, b: SeriesWithProv, *, units: str, transform: str, sources: list[SourceRef]) -> SeriesWithProv:
    years = sorted(set(a.values_by_year.keys()) | set(b.values_by_year.keys()))
    return SeriesWithProv(
        values_by_year={y: float(a.get(y)) + float(b.get(y)) for y in years},
        provenance=Provenance(tuple(sources), units=units, transform=transform),
    )


def build_baseline_energy_trust(
    parsed: RoadlightParsed,
    *,
    analysis_years: int = 10,
    discount_rate_annual: float = 0.12,
) -> BaselineResults:
    years = [y for y in parsed.years if 1 <= y <= int(analysis_years)]

    revenue_10k = parsed.baseline_revenue_trust_fee_rmb_y.reindex_years(years)
    cash_opex_10k = parsed.baseline_opex_cash_rmb_y.reindex_years(years)
    elec_opex_10k = parsed.baseline_opex_electricity_rmb_y.reindex_years(years)
    depr_10k = parsed.depreciation_rmb_y.reindex_years(years)

    revenue = to_rmb(revenue_10k)
    cash_opex = to_rmb(cash_opex_10k)
    elec_opex = to_rmb(elec_opex_10k)
    depr = to_rmb(depr_10k)

    accounting_gp = _series_sub(
        revenue,
        _series_add(
            cash_opex,
            depr,
            units="RMB_per_year",
            transform="cash_opex_plus_depreciation = cash_opex + depreciation.",
            sources=[
                *cash_opex.provenance.sources,
                *depr.provenance.sources,
            ],
        ),
        units="RMB_per_year",
        transform="accounting_gross_profit = revenue − (cash_opex + depreciation).",
        sources=[
            *revenue.provenance.sources,
            *cash_opex.provenance.sources,
            *depr.provenance.sources,
        ],
    )

    debt_p = to_rmb(parsed.debt_principal_rmb_y.reindex_years(years))
    debt_i = to_rmb(parsed.debt_interest_rmb_y.reindex_years(years))
    debt_service = _series_add(
        debt_p,
        debt_i,
        units="RMB_per_year",
        transform="debt_service = principal + interest.",
        sources=[*debt_p.provenance.sources, *debt_i.provenance.sources],
    )

    # CFADS definition (simple & traceable):
    # Use pre-debt operating cashflow: revenue − cash OPEX
    cfads = _series_sub(
        revenue,
        cash_opex,
        units="RMB_per_year",
        transform="CFADS = revenue − cash OPEX (simple, excludes depreciation).",
        sources=[*revenue.provenance.sources, *cash_opex.provenance.sources],
    )
    dscr = compute_dscr_by_year(
        cfads_rmb_y=cfads.values_by_year,
        debt_service_rmb_y=debt_service.values_by_year,
    )

    capex_y0 = float(parsed.capex_cash_rmb_y0)

    # Monthly timeline (month0 indexed):
    # month0 includes CAPEX outflow; months 1..N include equal monthly allocation of annual totals.
    n_months = int(analysis_years) * 12
    project_flows: list[float] = [-capex_y0]
    equity_flows: list[float] = [-capex_y0]  # debt inflow not modeled as equity inflow here; we model debt as outflow only for simplicity/traceability.

    # Expand annual series to monthly (flat within year).
    # This keeps payback definition aligned and explicit; refinement (seasonality) can come later.
    debt_service_monthly_by_year = {y: float(debt_service.get(y)) / 12.0 for y in years}
    op_cash_monthly_by_year = {y: float(cfads.get(y)) / 12.0 for y in years}

    for m in range(1, n_months + 1):
        y = (m - 1) // 12 + 1
        op = float(op_cash_monthly_by_year.get(y, 0.0))
        ds = float(debt_service_monthly_by_year.get(y, 0.0))
        project_flows.append(op)
        equity_flows.append(op - ds)

    payback = payback_month_from_monthly_cashflows(project_flows)
    irr_proj = irr_annual_from_monthly_cashflows(project_flows)
    irr_eq = irr_annual_from_monthly_cashflows(equity_flows)
    npv_proj = npv_monthly(project_flows, discount_rate_annual)
    npv_eq = npv_monthly(equity_flows, discount_rate_annual)

    return BaselineResults(
        years=years,
        revenue_rmb_y=revenue,
        cash_opex_rmb_y=cash_opex,
        electricity_opex_rmb_y=elec_opex,
        depreciation_rmb_y=depr,
        accounting_gross_profit_rmb_y=accounting_gp,
        debt_principal_rmb_y=debt_p,
        debt_interest_rmb_y=debt_i,
        debt_service_rmb_y=debt_service,
        dscr=dscr,
        project_cashflows_month0=project_flows,
        equity_cashflows_month0=equity_flows,
        payback_months=payback,
        irr_project_annual=irr_proj,
        irr_equity_annual=irr_eq,
        npv_project_rmb=float(npv_proj),
        npv_equity_rmb=float(npv_eq),
        capex_y0_rmb=capex_y0,
        capex_provenance=parsed.capex_provenance,
    )

