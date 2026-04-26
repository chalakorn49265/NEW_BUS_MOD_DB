from __future__ import annotations

from dataclasses import dataclass
from typing import Literal

import numpy as np

from Dashboard_LaaS_vs_NYTG.product_profiles import (
    ProductKey,
    forces_grid_electricity_zero,
    provenance_note as product_provenance_note,
    routine_om_scale,
)
from business_model_comparison.metrics import (
    DebtMetrics,
    compute_dscr_by_year,
    irr_annual_from_monthly_cashflows,
    npv_monthly,
    payback_month_from_monthly_cashflows,
)
from business_model_comparison.models import BaselineResults
from business_model_comparison.provenance import Provenance, SeriesWithProv, SourceRef

MAX_AI_OPEX_REDUCTION_PCT = 0.85
MIN_LAST_FOUR_YEAR_FEE_PCT = 0.20


@dataclass(frozen=True)
class LaaSScenario:
    term_years: int
    annual_service_fee_rmb: float  # client payment / provider revenue
    upfront_rmb: float  # client pays at month0 (positive to provider)
    ai_opex_reduction_pct: float  # fraction in [0, 1] (used when mode uses pct)
    last_four_year_fee_reduction_rmb: float = 0.0  # applied only in years 7-10; capped by floor
    opex_mode: Literal["uniform_pct", "electricity_only_pct", "ai_plus_solar"] = "uniform_pct"
    product_key: ProductKey = "AI_lightning_grid"


@dataclass(frozen=True)
class ClientValueAssumptions:
    """Assumptions that convert SLA/risk transfer into monetary value to the client."""

    baseline_outage_hours_per_year: float = 30.0
    laas_guaranteed_outage_hours_per_year: float = 5.0
    outage_cost_rmb_per_hour: float = 10_000.0
    sla_credit_share_to_client: float = 1.0  # 0..1
    client_discount_rate_annual: float = 0.12


@dataclass(frozen=True)
class LaaSResults:
    scenario: LaaSScenario
    years: list[int]

    client_payment_rmb_y: SeriesWithProv
    provider_revenue_rmb_y: SeriesWithProv
    provider_cash_opex_rmb_y: SeriesWithProv
    provider_depreciation_rmb_y: SeriesWithProv
    provider_accounting_gross_profit_rmb_y: SeriesWithProv

    debt_service_rmb_y: SeriesWithProv
    dscr: DebtMetrics

    project_cashflows_month0: list[float]
    payback_months: int | Literal["NO_PAYBACK"]
    irr_project_annual: float | Literal["NO_IRR"]
    npv_project_rmb: float

    # Constraint flags
    meets_pay_less_each_year: bool
    meets_provider_gross_profit_each_year: bool
    meets_payback_36m: bool
    meets_term_le_10: bool
    meets_payback_faster_than_baseline: bool
    client_benefit_pass: bool
    client_gap_rmb: float
    baseline_client_npv_cost_rmb: float
    laas_client_npv_cost_rmb: float
    guarantees_npv_value_rmb: float
    average_client_payment_rmb_per_year: float
    min_client_savings_rmb_per_year: float
    min_provider_gross_profit_uplift_rmb_per_year: float
    payback_improvement_months: float
    provider_feasible: bool
    feasible_everyone_better_off: bool

    # Traceability for “solved / enumerated”
    feasibility_provenance: Provenance


def _flat_annual_series(years: list[int], value_rmb_per_year: float, *, file_note: str, label: str, transform: str) -> SeriesWithProv:
    return SeriesWithProv(
        values_by_year={int(y): float(value_rmb_per_year) for y in years},
        provenance=Provenance(
            sources=(SourceRef(file=file_note, row_label=label),),
            units="RMB_per_year",
            transform=transform,
        ),
    )


def _build_client_payment_schedule(
    *,
    years: list[int],
    annual_fee_rmb: float,
    last_four_year_fee_reduction_rmb: float,
    upfront_rmb: float,
) -> SeriesWithProv:
    annual_fee = max(0.0, float(annual_fee_rmb))
    tail_reduction = max(0.0, float(last_four_year_fee_reduction_rmb))
    prepaid_per_year = max(0.0, float(upfront_rmb)) / max(1, len(years))
    tail_fee_floor = annual_fee * MIN_LAST_FOUR_YEAR_FEE_PCT

    values_by_year: dict[int, float] = {}
    for y in years:
        gross_fee_y = annual_fee
        if y >= 7:
            gross_fee_y = max(tail_fee_floor, annual_fee - tail_reduction)
        values_by_year[int(y)] = max(0.0, gross_fee_y - prepaid_per_year)

    return SeriesWithProv(
        values_by_year=values_by_year,
        provenance=Provenance(
            sources=(
                SourceRef(file="ASSUMPTION/COMPUTED", row_label="LaaS_annual_service_fee_rmb", notes=f"{annual_fee:.2f}"),
                SourceRef(file="ASSUMPTION/COMPUTED", row_label="last_four_year_fee_reduction_rmb", notes=f"{tail_reduction:.2f}"),
                SourceRef(file="ASSUMPTION/COMPUTED", row_label="upfront_rmb", notes=f"{float(upfront_rmb):.2f}"),
                SourceRef(file="ASSUMPTION/COMPUTED", row_label="tail_fee_floor_pct", notes=f"{MIN_LAST_FOUR_YEAR_FEE_PCT:.0%}"),
            ),
            units="RMB_per_year",
            transform=(
                "Client annual payment schedule: years 1-6 use annual_service_fee_rmb; years 7-10 use "
                "max(annual_service_fee_rmb × tail_fee_floor_pct, annual_service_fee_rmb − last_four_year_fee_reduction_rmb). "
                "Upfront is treated as prepayment and allocated evenly across term years."
            ),
        ),
    )


def _apply_uniform_reduction(cash_opex: SeriesWithProv, reduction_pct: float) -> SeriesWithProv:
    r = max(0.0, min(MAX_AI_OPEX_REDUCTION_PCT, float(reduction_pct)))
    return SeriesWithProv(
        values_by_year={y: float(v) * (1.0 - r) for y, v in cash_opex.values_by_year.items()},
        provenance=Provenance(
            sources=(
                *cash_opex.provenance.sources,
                SourceRef(file="ASSUMPTION/COMPUTED", row_label="opex_mode", notes="uniform_pct"),
                SourceRef(file="ASSUMPTION/COMPUTED", row_label="ai_opex_reduction_pct", notes=f"{r:.3f} applied to total cash OPEX"),
            ),
            units=cash_opex.provenance.units,
            transform=cash_opex.provenance.transform + f" Then multiplied by (1 − {r:.3f}) under uniform_pct mode.",
        ),
    )


def _apply_opex_mode(
    *,
    baseline_cash_opex: SeriesWithProv,
    baseline_electricity_opex: SeriesWithProv,
    mode: Literal["uniform_pct", "electricity_only_pct", "ai_plus_solar"],
    reduction_pct: float,
) -> SeriesWithProv:
    """Return provider cash OPEX series under selected mode, with provenance."""
    years = sorted(set(baseline_cash_opex.values_by_year.keys()) | set(baseline_electricity_opex.values_by_year.keys()))
    cash = baseline_cash_opex.reindex_years(years)
    elec = baseline_electricity_opex.reindex_years(years)
    other = SeriesWithProv(
        values_by_year={y: float(cash.get(y)) - float(elec.get(y)) for y in years},
        provenance=Provenance(
            sources=(
                *cash.provenance.sources,
                *elec.provenance.sources,
            ),
            units=cash.provenance.units,
            transform="other_cash_opex = total_cash_opex − electricity_opex.",
        ),
    )

    if mode == "uniform_pct":
        return _apply_uniform_reduction(cash, reduction_pct)

    if mode == "electricity_only_pct":
        r = max(0.0, min(MAX_AI_OPEX_REDUCTION_PCT, float(reduction_pct)))
        elec_adj = SeriesWithProv(
            values_by_year={y: float(elec.get(y)) * (1.0 - r) for y in years},
            provenance=Provenance(
                sources=(
                    *elec.provenance.sources,
                    SourceRef(file="ASSUMPTION/COMPUTED", row_label="opex_mode", notes="electricity_only_pct"),
                    SourceRef(file="ASSUMPTION/COMPUTED", row_label="ai_opex_reduction_pct", notes=f"{r:.3f} applied to electricity only"),
                ),
                units=elec.provenance.units,
                transform=elec.provenance.transform + f" Then multiplied by (1 − {r:.3f}) under electricity_only_pct mode.",
            ),
        )
        return SeriesWithProv(
            values_by_year={y: float(other.get(y)) + float(elec_adj.get(y)) for y in years},
            provenance=Provenance(
                sources=(*other.provenance.sources, *elec_adj.provenance.sources),
                units=elec.provenance.units,
                transform="cash_opex = other_cash_opex + adjusted_electricity_opex (electricity_only_pct).",
            ),
        )

    # ai_plus_solar: electricity becomes 0 for all years (no solar CAPEX modeled; must be explicit).
    if mode == "ai_plus_solar":
        elec_zero = SeriesWithProv(
            values_by_year={y: 0.0 for y in years},
            provenance=Provenance(
                sources=(
                    *elec.provenance.sources,
                    SourceRef(file="ASSUMPTION/COMPUTED", row_label="opex_mode", notes="ai_plus_solar"),
                    SourceRef(file="ASSUMPTION/COMPUTED", row_label="electricity_opex_rule", notes="Set electricity OPEX to 0 for all years; no solar CAPEX modeled."),
                ),
                units=elec.provenance.units,
                transform="Electricity OPEX set to 0 for all years under ai_plus_solar mode (assumption; no solar CAPEX modeled).",
            ),
        )
        return SeriesWithProv(
            values_by_year={y: float(other.get(y)) + float(elec_zero.get(y)) for y in years},
            provenance=Provenance(
                sources=(*other.provenance.sources, *elec_zero.provenance.sources),
                units=elec.provenance.units,
                transform="cash_opex = other_cash_opex + 0 electricity (ai_plus_solar).",
            ),
        )

    # Defensive fallback
    return cash


def _apply_opex_mode_split(
    *,
    baseline_cash_opex: SeriesWithProv,
    baseline_electricity_opex: SeriesWithProv,
    mode: Literal["uniform_pct", "electricity_only_pct", "ai_plus_solar"],
    reduction_pct: float,
) -> tuple[SeriesWithProv, SeriesWithProv, SeriesWithProv]:
    """
    Return (other_cash_opex, electricity_opex, total_cash_opex) after applying `mode`.

    This is used so product overlays (solar/battery) can be applied cleanly.
    """
    years = sorted(set(baseline_cash_opex.values_by_year.keys()) | set(baseline_electricity_opex.values_by_year.keys()))
    cash = baseline_cash_opex.reindex_years(years)
    elec = baseline_electricity_opex.reindex_years(years)
    other = SeriesWithProv(
        values_by_year={y: float(cash.get(y)) - float(elec.get(y)) for y in years},
        provenance=Provenance(
            sources=(
                *cash.provenance.sources,
                *elec.provenance.sources,
            ),
            units=cash.provenance.units,
            transform="other_cash_opex = total_cash_opex − electricity_opex.",
        ),
    )

    if mode == "uniform_pct":
        r = max(0.0, min(MAX_AI_OPEX_REDUCTION_PCT, float(reduction_pct)))
        other_adj = SeriesWithProv(
            values_by_year={y: float(other.get(y)) * (1.0 - r) for y in years},
            provenance=Provenance(
                sources=(
                    *other.provenance.sources,
                    SourceRef(file="ASSUMPTION/COMPUTED", row_label="opex_mode", notes="uniform_pct"),
                    SourceRef(file="ASSUMPTION/COMPUTED", row_label="ai_opex_reduction_pct", notes=f"{r:.3f} applied to non-electric portion"),
                ),
                units=other.provenance.units,
                transform=f"other_opex = other_cash_opex × (1 − {r:.3f}) under uniform_pct.",
            ),
        )
        elec_adj = SeriesWithProv(
            values_by_year={y: float(elec.get(y)) * (1.0 - r) for y in years},
            provenance=Provenance(
                sources=(
                    *elec.provenance.sources,
                    SourceRef(file="ASSUMPTION/COMPUTED", row_label="opex_mode", notes="uniform_pct"),
                    SourceRef(file="ASSUMPTION/COMPUTED", row_label="ai_opex_reduction_pct", notes=f"{r:.3f} applied to electricity portion"),
                ),
                units=elec.provenance.units,
                transform=f"electricity_opex = electricity_opex × (1 − {r:.3f}) under uniform_pct.",
            ),
        )
        cash_adj = SeriesWithProv(
            values_by_year={y: float(other_adj.get(y)) + float(elec_adj.get(y)) for y in years},
            provenance=Provenance(
                sources=(*other_adj.provenance.sources, *elec_adj.provenance.sources),
                units=cash.provenance.units,
                transform="cash_opex = other_opex + electricity_opex (uniform_pct split form).",
            ),
        )
        return other_adj, elec_adj, cash_adj

    if mode == "electricity_only_pct":
        r = max(0.0, min(MAX_AI_OPEX_REDUCTION_PCT, float(reduction_pct)))
        elec_adj = SeriesWithProv(
            values_by_year={y: float(elec.get(y)) * (1.0 - r) for y in years},
            provenance=Provenance(
                sources=(
                    *elec.provenance.sources,
                    SourceRef(file="ASSUMPTION/COMPUTED", row_label="opex_mode", notes="electricity_only_pct"),
                    SourceRef(file="ASSUMPTION/COMPUTED", row_label="ai_opex_reduction_pct", notes=f"{r:.3f} applied to electricity only"),
                ),
                units=elec.provenance.units,
                transform=elec.provenance.transform + f" Then multiplied by (1 − {r:.3f}) under electricity_only_pct mode.",
            ),
        )
        cash_adj = SeriesWithProv(
            values_by_year={y: float(other.get(y)) + float(elec_adj.get(y)) for y in years},
            provenance=Provenance(
                sources=(*other.provenance.sources, *elec_adj.provenance.sources),
                units=elec.provenance.units,
                transform="cash_opex = other_cash_opex + adjusted_electricity_opex (electricity_only_pct).",
            ),
        )
        return other, elec_adj, cash_adj

    if mode == "ai_plus_solar":
        elec_zero = SeriesWithProv(
            values_by_year={y: 0.0 for y in years},
            provenance=Provenance(
                sources=(
                    *elec.provenance.sources,
                    SourceRef(file="ASSUMPTION/COMPUTED", row_label="opex_mode", notes="ai_plus_solar"),
                    SourceRef(file="ASSUMPTION/COMPUTED", row_label="electricity_opex_rule", notes="Set electricity OPEX to 0 for all years; no solar CAPEX modeled."),
                ),
                units=elec.provenance.units,
                transform="Electricity OPEX set to 0 for all years under ai_plus_solar mode (assumption; no solar CAPEX modeled).",
            ),
        )
        cash_adj = SeriesWithProv(
            values_by_year={y: float(other.get(y)) + float(elec_zero.get(y)) for y in years},
            provenance=Provenance(
                sources=(*other.provenance.sources, *elec_zero.provenance.sources),
                units=elec.provenance.units,
                transform="cash_opex = other_cash_opex + 0 electricity (ai_plus_solar).",
            ),
        )
        return other, elec_zero, cash_adj

    return other, elec, cash

def _client_value_from_guarantees_by_year(
    years: list[int],
    a: ClientValueAssumptions,
) -> SeriesWithProv:
    base_h = max(0.0, float(a.baseline_outage_hours_per_year))
    laas_h = max(0.0, float(a.laas_guaranteed_outage_hours_per_year))
    delta_h = max(0.0, base_h - laas_h)
    cost = max(0.0, float(a.outage_cost_rmb_per_hour))
    share = max(0.0, min(1.0, float(a.sla_credit_share_to_client)))
    v_y = delta_h * cost * share
    return SeriesWithProv(
        values_by_year={int(y): float(v_y) for y in years},
        provenance=Provenance(
            sources=(
                SourceRef(file="ASSUMPTION/COMPUTED", row_label="baseline_outage_hours_per_year", notes=str(base_h)),
                SourceRef(file="ASSUMPTION/COMPUTED", row_label="laas_guaranteed_outage_hours_per_year", notes=str(laas_h)),
                SourceRef(file="ASSUMPTION/COMPUTED", row_label="outage_cost_rmb_per_hour", notes=str(cost)),
                SourceRef(file="ASSUMPTION/COMPUTED", row_label="sla_credit_share_to_client", notes=str(share)),
            ),
            units="RMB_per_year",
            transform="ValueFromGuarantees_y = max(0, baseline_outage_hours − laas_outage_hours) × outage_cost × credit_share.",
        ),
    )


def _pv_cost_from_annual_payments(
    *,
    years: list[int],
    annual_payment_rmb_y: SeriesWithProv,
    upfront_rmb_month0: float,
    annual_discount_rate: float,
) -> float:
    n_months = int(max(years) if years else 0) * 12
    flows: list[float] = [float(upfront_rmb_month0)]
    for m in range(1, n_months + 1):
        y = (m - 1) // 12 + 1
        p = float(annual_payment_rmb_y.get(y, 0.0)) / 12.0
        flows.append(float(p))
    # NPV of costs (positive outflows) discounted from month0.
    return float(npv_monthly(flows, float(annual_discount_rate)))


def _pv_value_from_annual_benefits(
    *,
    years: list[int],
    annual_benefit_rmb_y: SeriesWithProv,
    annual_discount_rate: float,
) -> float:
    n_months = int(max(years) if years else 0) * 12
    flows: list[float] = [0.0]
    for m in range(1, n_months + 1):
        y = (m - 1) // 12 + 1
        b = float(annual_benefit_rmb_y.get(y, 0.0)) / 12.0
        flows.append(float(b))
    return float(npv_monthly(flows, float(annual_discount_rate)))


def evaluate_laas_scenario(
    baseline: BaselineResults,
    scenario: LaaSScenario,
    *,
    discount_rate_annual: float,
    client_value: ClientValueAssumptions | None = None,
) -> LaaSResults:
    term = int(scenario.term_years)
    years = [y for y in baseline.years if 1 <= y <= term]

    # Commercial convention: upfront is treated as a prepayment that reduces subsequent annual fees.
    # We allocate upfront evenly across term years (transparent and easy to audit).
    annual_fee = float(scenario.annual_service_fee_rmb)
    upfront = float(scenario.upfront_rmb)
    client_pay = _build_client_payment_schedule(
        years=years,
        annual_fee_rmb=annual_fee,
        last_four_year_fee_reduction_rmb=float(scenario.last_four_year_fee_reduction_rmb),
        upfront_rmb=upfront,
    )
    provider_rev = SeriesWithProv(
        values_by_year=dict(client_pay.values_by_year),
        provenance=Provenance(
            sources=client_pay.provenance.sources,
            units="RMB_per_year",
            transform="Provider revenue under LaaS assumed equal to client service fee payments (same net-of-prepayment series).",
        ),
    )

    base_cash = baseline.cash_opex_rmb_y.reindex_years(years)
    base_elec = baseline.electricity_opex_rmb_y.reindex_years(years)
    other_mode, elec_mode, _cash_mode = _apply_opex_mode_split(
        baseline_cash_opex=base_cash,
        baseline_electricity_opex=base_elec,
        mode=scenario.opex_mode,
        reduction_pct=scenario.ai_opex_reduction_pct,
    )

    # Product overlay: solar forces electricity to zero; battery reduces routine (non-electric) OPEX.
    if forces_grid_electricity_zero(scenario.product_key):
        elec_adj = SeriesWithProv(
            values_by_year={y: 0.0 for y in years},
            provenance=Provenance(
                sources=(
                    *elec_mode.provenance.sources,
                    SourceRef(file="ASSUMPTION/COMPUTED", row_label="product_key", notes=str(scenario.product_key)),
                    SourceRef(file="ASSUMPTION/COMPUTED", row_label="electricity_opex_rule", notes="Off-grid solar product: electricity OPEX set to 0."),
                ),
                units=elec_mode.provenance.units,
                transform="electricity_opex = 0 for all years due to off-grid solar product (no grid purchases).",
            ),
        )
    else:
        elec_adj = elec_mode

    om_scale = routine_om_scale(scenario.product_key)
    other_scaled = SeriesWithProv(
        values_by_year={y: max(0.0, float(other_mode.get(y)) * om_scale) for y in years},
        provenance=Provenance(
            sources=(
                *other_mode.provenance.sources,
                SourceRef(file="ASSUMPTION/COMPUTED", row_label="product_key", notes=str(scenario.product_key)),
                SourceRef(file="ASSUMPTION/COMPUTED", row_label="routine_om_scale", notes=f"{om_scale:.3f}; {product_provenance_note(scenario.product_key)}"),
            ),
            units=base_cash.provenance.units,
            transform=f"non_electric_opex_scaled = max(0, other_opex_after_mode × {om_scale:.3f}) based on product routine_maint factor.",
        ),
    )
    cash_opex = SeriesWithProv(
        values_by_year={y: float(other_scaled.get(y)) + float(elec_adj.get(y)) for y in years},
        provenance=Provenance(
            sources=(*other_scaled.provenance.sources, *elec_adj.provenance.sources),
            units=base_cash.provenance.units,
            transform="cash_opex = non_electric_opex_scaled + electricity_opex (mode + product overlay).",
        ),
    )
    depr = baseline.depreciation_rmb_y.reindex_years(years)  # accounting depreciation (kept same unless later refined)

    gp = SeriesWithProv(
        values_by_year={y: float(provider_rev.get(y)) - (float(cash_opex.get(y)) + float(depr.get(y))) for y in years},
        provenance=Provenance(
            sources=(
                *provider_rev.provenance.sources,
                *cash_opex.provenance.sources,
                *depr.provenance.sources,
            ),
            units="RMB_per_year",
            transform="accounting_gross_profit = revenue − (cash OPEX + depreciation).",
        ),
    )

    # Debt stays the same schedule for now (same project financing assumption).
    debt_service = baseline.debt_service_rmb_y.reindex_years(years)

    # DSCR uses CFADS = revenue − cash OPEX
    cfads = {y: float(provider_rev.get(y)) - float(cash_opex.get(y)) for y in years}
    dscr = compute_dscr_by_year(cfads_rmb_y=cfads, debt_service_rmb_y=debt_service.values_by_year)

    # Monthly cashflows: month0 includes -CAPEX + upfront inflow; months use net annual fee schedule.
    n_months = term * 12
    capex = float(baseline.capex_y0_rmb)  # keep same equipment CAPEX (source: capex.csv)
    project_flows: list[float] = [-capex + upfront]
    for m in range(1, n_months + 1):
        y = (m - 1) // 12 + 1
        rev_m = float(provider_rev.get(y, 0.0)) / 12.0
        opex_m = float(cash_opex.get(y, 0.0)) / 12.0
        project_flows.append(rev_m - opex_m)

    payback = payback_month_from_monthly_cashflows(project_flows)
    irr_p = irr_annual_from_monthly_cashflows(project_flows)
    npv_p = npv_monthly(project_flows, float(discount_rate_annual))

    # Client value model (PV-based, traceable)
    cv = client_value or ClientValueAssumptions(client_discount_rate_annual=float(discount_rate_annual))
    baseline_client_payment = baseline.revenue_rmb_y.reindex_years(years)
    baseline_pv_cost = _pv_cost_from_annual_payments(
        years=years,
        annual_payment_rmb_y=baseline_client_payment,
        upfront_rmb_month0=0.0,
        annual_discount_rate=float(cv.client_discount_rate_annual),
    )
    laas_pv_cost = _pv_cost_from_annual_payments(
        years=years,
        annual_payment_rmb_y=client_pay,
        upfront_rmb_month0=float(scenario.upfront_rmb),
        annual_discount_rate=float(cv.client_discount_rate_annual),
    )
    guarantees_y = _client_value_from_guarantees_by_year(years, cv)
    guarantees_pv = _pv_value_from_annual_benefits(
        years=years,
        annual_benefit_rmb_y=guarantees_y,
        annual_discount_rate=float(cv.client_discount_rate_annual),
    )
    # client_gap > 0 means client worse off after accounting for guarantee value.
    client_gap = float(laas_pv_cost - baseline_pv_cost - guarantees_pv)
    client_benefit_pass = client_gap <= 0.0 + 1e-6
    average_client_payment = (
        sum(float(client_pay.get(y)) for y in years) / len(years)
        if years
        else 0.0
    )

    min_client_savings = min(float(baseline.revenue_rmb_y.get(y)) - float(client_pay.get(y)) for y in years) if years else 0.0
    min_provider_gp_uplift = min(
        float(gp.get(y)) - float(baseline.accounting_gross_profit_rmb_y.get(y))
        for y in years
    ) if years else 0.0

    baseline_payback = baseline.payback_months
    payback_faster_than_baseline = (
        isinstance(payback, int)
        and isinstance(baseline_payback, int)
        and payback < baseline_payback
    )
    payback_improvement_months = (
        float(baseline_payback - payback)
        if isinstance(payback, int) and isinstance(baseline_payback, int)
        else float("nan")
    )

    # Constraints
    pays_less_each_year = all(float(client_pay.get(y)) <= float(baseline.revenue_rmb_y.get(y)) + 1e-6 for y in years)
    provider_gp_each_year = all(float(gp.get(y)) >= float(baseline.accounting_gross_profit_rmb_y.get(y)) - 1e-6 for y in years)
    payback_ok = isinstance(payback, int) and payback <= 36
    term_ok = term <= 10
    provider_feasible = provider_gp_each_year and payback_ok and term_ok
    feasible = provider_feasible and pays_less_each_year and client_benefit_pass and payback_faster_than_baseline

    feas_prov = Provenance(
        sources=(
            SourceRef(
                file="ASSUMPTION/COMPUTED",
                row_label="feasible_region",
                notes="Evaluated via explicit grid search over (term, fee, last-4-year reduction, upfront, OPEX mode, AI reduction, and client value assumptions).",
            ),
            *baseline.capex_provenance.sources,
            *baseline.revenue_rmb_y.provenance.sources,
            *baseline.cash_opex_rmb_y.provenance.sources,
            *baseline.electricity_opex_rmb_y.provenance.sources,
            *baseline.depreciation_rmb_y.provenance.sources,
            *baseline.debt_service_rmb_y.provenance.sources,
        ),
        units="boolean",
        transform=(
            "Provider-feasible iff: term<=10 AND payback<=36m AND provider_gross_profit>=baseline (each year). "
            "Everyone-feasible iff: provider-feasible AND pays_less_each_year AND payback improves vs baseline AND client_gap<=0, "
            "where client_gap = PV(LaaS payments incl upfront) − PV(baseline payments) − PV(ValueFromGuarantees)."
        ),
    )

    return LaaSResults(
        scenario=scenario,
        years=years,
        client_payment_rmb_y=client_pay,
        provider_revenue_rmb_y=provider_rev,
        provider_cash_opex_rmb_y=cash_opex,
        provider_depreciation_rmb_y=depr,
        provider_accounting_gross_profit_rmb_y=gp,
        debt_service_rmb_y=debt_service,
        dscr=dscr,
        project_cashflows_month0=project_flows,
        payback_months=payback,
        irr_project_annual=irr_p,
        npv_project_rmb=float(npv_p),
        meets_pay_less_each_year=bool(pays_less_each_year),
        meets_provider_gross_profit_each_year=bool(provider_gp_each_year),
        meets_payback_36m=bool(payback_ok),
        meets_term_le_10=bool(term_ok),
        meets_payback_faster_than_baseline=bool(payback_faster_than_baseline),
        client_benefit_pass=bool(client_benefit_pass),
        client_gap_rmb=float(client_gap),
        baseline_client_npv_cost_rmb=float(baseline_pv_cost),
        laas_client_npv_cost_rmb=float(laas_pv_cost),
        guarantees_npv_value_rmb=float(guarantees_pv),
        average_client_payment_rmb_per_year=float(average_client_payment),
        min_client_savings_rmb_per_year=float(min_client_savings),
        min_provider_gross_profit_uplift_rmb_per_year=float(min_provider_gp_uplift),
        payback_improvement_months=float(payback_improvement_months),
        provider_feasible=bool(provider_feasible),
        feasible_everyone_better_off=bool(feasible),
        feasibility_provenance=feas_prov,
    )


@dataclass(frozen=True)
class FeasibleEnvelope:
    """Aggregated feasible region results for plotting and offer boundaries."""

    evaluated: int
    feasible: int
    results: list[LaaSResults]

    def best_by_provider_npv(self) -> LaaSResults | None:
        feas = [r for r in self.results if r.feasible_everyone_better_off]
        if not feas:
            return None
        return max(feas, key=lambda r: float(r.npv_project_rmb))


def grid_search_feasible_envelope(
    baseline: BaselineResults,
    *,
    term_years: list[int],
    annual_fee_rmb_grid: list[float],
    last_four_year_fee_reduction_rmb_grid: list[float] | None,
    upfront_rmb_grid: list[float],
    ai_opex_reduction_grid: list[float],
    discount_rate_annual: float,
    opex_modes: list[Literal["uniform_pct", "electricity_only_pct", "ai_plus_solar"]] | None = None,
    client_value: ClientValueAssumptions | None = None,
) -> FeasibleEnvelope:
    out: list[LaaSResults] = []
    evaluated = 0
    modes = opex_modes or ["uniform_pct"]
    tail_reduction_grid = sorted({max(0.0, float(x)) for x in (last_four_year_fee_reduction_rmb_grid or [0.0])})
    normalized_reduction_grid = sorted(
        {
            max(0.0, min(MAX_AI_OPEX_REDUCTION_PCT, float(red)))
            for red in ai_opex_reduction_grid
        }
    )
    for term in term_years:
        applicable_tail_reduction_grid = tail_reduction_grid if int(term) > 6 else [0.0]
        for fee in annual_fee_rmb_grid:
            for tail_red in applicable_tail_reduction_grid:
                for up in upfront_rmb_grid:
                    for red in normalized_reduction_grid:
                        for mode in modes:
                            evaluated += 1
                            scen = LaaSScenario(
                                term_years=int(term),
                                annual_service_fee_rmb=float(fee),
                                last_four_year_fee_reduction_rmb=float(tail_red),
                                upfront_rmb=float(up),
                                ai_opex_reduction_pct=float(red),
                                opex_mode=mode,
                            )
                            out.append(
                                evaluate_laas_scenario(
                                    baseline,
                                    scen,
                                    discount_rate_annual=float(discount_rate_annual),
                                    client_value=client_value,
                                )
                            )
    feasible = sum(1 for r in out if r.feasible_everyone_better_off)
    return FeasibleEnvelope(evaluated=evaluated, feasible=feasible, results=out)


def default_fee_grid_from_baseline(baseline: BaselineResults, *, pct_low: float = 0.4, pct_high: float = 1.2, steps: int = 60) -> list[float]:
    # Baseline fee is flat in your dataset; take year-1 as reference.
    base = float(baseline.revenue_rmb_y.get(1))
    lo = max(0.0, base * float(pct_low))
    hi = max(lo, base * float(pct_high))
    return [float(x) for x in np.linspace(lo, hi, int(steps))]

