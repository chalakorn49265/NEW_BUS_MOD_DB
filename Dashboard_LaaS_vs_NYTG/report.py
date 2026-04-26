from __future__ import annotations

from dataclasses import asdict

import pandas as pd

from business_model_comparison.laas_feasible import FeasibleEnvelope, LaaSResults
from business_model_comparison.models import BaselineResults


def baseline_summary_table(b: BaselineResults) -> pd.DataFrame:
    rows = []
    cumulative = 0.0
    year0_net = -float(b.capex_y0_rmb)
    cumulative += year0_net
    rows.append(
        {
            "year": 0,
            "capex_rmb": float(b.capex_y0_rmb),
            "service_fee_rmb": 0.0,
            "upfront_rmb": 0.0,
            "net_cashflow_cumulative_rmb": cumulative,
        }
    )
    for y in b.years:
        cumulative += float(b.revenue_rmb_y.get(y)) - float(b.cash_opex_rmb_y.get(y))
        rows.append(
            {
                "year": y,
                "capex_rmb": 0.0,
                "service_fee_rmb": float(b.revenue_rmb_y.get(y)),
                "upfront_rmb": 0.0,
                "net_cashflow_cumulative_rmb": cumulative,
            }
        )
    return pd.DataFrame(rows)


def laas_results_to_table(r: LaaSResults, baseline: BaselineResults) -> pd.DataFrame:
    years = r.years
    rows = []
    rows.append(
        {
            "year": 0,
            "baseline_capex_rmb": float(baseline.capex_y0_rmb),
            "laas_capex_rmb": float(baseline.capex_y0_rmb),
            "baseline_trust_fee_rmb": 0.0,
            "provider_revenue_rmb": 0.0,
            "upfront_payment_rmb": float(r.scenario.upfront_rmb),
            "laas_service_fee_rmb": 0.0,
            "total_client_outflow_rmb": float(r.scenario.upfront_rmb),
            "client_savings_rmb": -float(r.scenario.upfront_rmb),
            "baseline_cash_opex_rmb": 0.0,
            "baseline_electricity_opex_rmb": 0.0,
            "laas_cash_opex_rmb": 0.0,
            "baseline_depreciation_rmb": 0.0,
            "laas_depreciation_rmb": 0.0,
            "baseline_gross_profit_rmb": 0.0,
            "laas_gross_profit_rmb": 0.0,
            "gross_profit_uplift_rmb": 0.0,
            "baseline_project_cashflow_rmb": -float(baseline.capex_y0_rmb),
            "laas_project_cashflow_rmb": -float(baseline.capex_y0_rmb) + float(r.scenario.upfront_rmb),
            "debt_principal_rmb": 0.0,
            "debt_interest_rmb": 0.0,
            "debt_service_rmb": 0.0,
            "dscr": None,
        }
    )
    for y in years:
        laas_fee = float(r.client_payment_rmb_y.get(y))
        provider_revenue = float(r.provider_revenue_rmb_y.get(y))
        laas_cash_opex = float(r.provider_cash_opex_rmb_y.get(y))
        laas_depreciation = float(r.provider_depreciation_rmb_y.get(y))
        baseline_revenue = float(baseline.revenue_rmb_y.get(y))
        baseline_cash_opex = float(baseline.cash_opex_rmb_y.get(y))
        rows.append(
            {
                "year": y,
                "baseline_capex_rmb": 0.0,
                "laas_capex_rmb": 0.0,
                "baseline_trust_fee_rmb": baseline_revenue,
                "provider_revenue_rmb": provider_revenue,
                "upfront_payment_rmb": 0.0,
                "laas_service_fee_rmb": laas_fee,
                "total_client_outflow_rmb": laas_fee,
                "client_savings_rmb": baseline_revenue - laas_fee,
                "baseline_cash_opex_rmb": baseline_cash_opex,
                "baseline_electricity_opex_rmb": float(baseline.electricity_opex_rmb_y.get(y)),
                "laas_cash_opex_rmb": laas_cash_opex,
                "baseline_depreciation_rmb": float(baseline.depreciation_rmb_y.get(y)),
                "laas_depreciation_rmb": laas_depreciation,
                "baseline_gross_profit_rmb": float(baseline.accounting_gross_profit_rmb_y.get(y)),
                "laas_gross_profit_rmb": float(r.provider_accounting_gross_profit_rmb_y.get(y)),
                "gross_profit_uplift_rmb": float(r.provider_accounting_gross_profit_rmb_y.get(y))
                - float(baseline.accounting_gross_profit_rmb_y.get(y)),
                "baseline_project_cashflow_rmb": baseline_revenue - baseline_cash_opex,
                "laas_project_cashflow_rmb": provider_revenue - laas_cash_opex,
                "debt_principal_rmb": float(baseline.debt_principal_rmb_y.get(y)),
                "debt_interest_rmb": float(baseline.debt_interest_rmb_y.get(y)),
                "debt_service_rmb": float(r.debt_service_rmb_y.get(y)),
                "dscr": (None if r.dscr.dscr_by_year.get(y) is None else float(r.dscr.dscr_by_year[y])),
            }
        )
    return pd.DataFrame(rows)


def envelope_table(env: FeasibleEnvelope) -> pd.DataFrame:
    rows = []
    for r in env.results:
        s = r.scenario
        rows.append(
            {
                "term_years": int(s.term_years),
                "annual_service_fee_rmb": float(s.annual_service_fee_rmb),
                "last_four_year_fee_reduction_rmb": float(s.last_four_year_fee_reduction_rmb),
                "upfront_rmb": float(s.upfront_rmb),
                "ai_opex_reduction_pct": float(s.ai_opex_reduction_pct),
                "opex_mode": str(s.opex_mode),
                "payback_months": r.payback_months if isinstance(r.payback_months, str) else int(r.payback_months),
                "irr_project_annual": r.irr_project_annual if isinstance(r.irr_project_annual, str) else float(r.irr_project_annual),
                "npv_project_rmb": float(r.npv_project_rmb),
                "dscr_min": (None if r.dscr.dscr_min is None else float(r.dscr.dscr_min)),
                "meets_pay_less_each_year": bool(r.meets_pay_less_each_year),
                "meets_provider_gross_profit_each_year": bool(r.meets_provider_gross_profit_each_year),
                "meets_payback_36m": bool(r.meets_payback_36m),
                "meets_payback_faster_than_baseline": bool(r.meets_payback_faster_than_baseline),
                "provider_feasible": bool(r.provider_feasible),
                "client_benefit_pass": bool(r.client_benefit_pass),
                "client_gap_rmb": float(r.client_gap_rmb),
                "baseline_client_npv_cost_rmb": float(r.baseline_client_npv_cost_rmb),
                "laas_client_npv_cost_rmb": float(r.laas_client_npv_cost_rmb),
                "guarantees_npv_value_rmb": float(r.guarantees_npv_value_rmb),
                "average_client_payment_rmb_per_year": float(r.average_client_payment_rmb_per_year),
                "min_client_savings_rmb_per_year": float(r.min_client_savings_rmb_per_year),
                "min_provider_gross_profit_uplift_rmb_per_year": float(r.min_provider_gross_profit_uplift_rmb_per_year),
                "payback_improvement_months": float(r.payback_improvement_months),
                "feasible_everyone_better_off": bool(r.feasible_everyone_better_off),
            }
        )
    return pd.DataFrame(rows)


def simple_cashflow_comparison_table(r: LaaSResults, baseline: BaselineResults) -> pd.DataFrame:
    rows = []
    baseline_cum = 0.0
    laas_cum = 0.0

    baseline_y0 = -float(baseline.capex_y0_rmb)
    laas_y0 = -float(baseline.capex_y0_rmb) + float(r.scenario.upfront_rmb)
    baseline_cum += baseline_y0
    laas_cum += laas_y0
    rows.append(
        {
            "year": 0,
            "trust_capex_rmb": float(baseline.capex_y0_rmb),
            "trust_service_fee_rmb": 0.0,
            "trust_upfront_rmb": 0.0,
            "trust_cash_opex_rmb": 0.0,
            "trust_net_cashflow_rmb": float(baseline_y0),
            "trust_net_cashflow_cumulative_rmb": baseline_cum,
            "laas_capex_rmb": float(baseline.capex_y0_rmb),
            "laas_service_fee_rmb": 0.0,
            "laas_upfront_rmb": float(r.scenario.upfront_rmb),
            "laas_cash_opex_rmb": 0.0,
            "laas_net_cashflow_rmb": float(laas_y0),
            "laas_net_cashflow_cumulative_rmb": laas_cum,
        }
    )

    for y in baseline.years:
        trust_fee = float(baseline.revenue_rmb_y.get(y))
        trust_opex = float(baseline.cash_opex_rmb_y.get(y))
        baseline_net = trust_fee - trust_opex

        laas_fee = float(r.client_payment_rmb_y.get(y))
        laas_opex = float(r.provider_cash_opex_rmb_y.get(y))
        laas_net = float(r.provider_revenue_rmb_y.get(y)) - laas_opex

        baseline_cum += baseline_net
        laas_cum += laas_net
        rows.append(
            {
                "year": int(y),
                "trust_capex_rmb": 0.0,
                "trust_service_fee_rmb": trust_fee,
                "trust_upfront_rmb": 0.0,
                "trust_cash_opex_rmb": trust_opex,
                "trust_net_cashflow_rmb": baseline_net,
                "trust_net_cashflow_cumulative_rmb": baseline_cum,
                "laas_capex_rmb": 0.0,
                "laas_service_fee_rmb": laas_fee,
                "laas_upfront_rmb": 0.0,
                "laas_cash_opex_rmb": laas_opex,
                "laas_net_cashflow_rmb": laas_net,
                "laas_net_cashflow_cumulative_rmb": laas_cum,
            }
        )
    return pd.DataFrame(rows)


def rank_recommended_offers(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return df
    sort_cols = [
        "min_client_savings_rmb_per_year",
        "min_provider_gross_profit_uplift_rmb_per_year",
        "payback_months",
        "npv_project_rmb",
        "last_four_year_fee_reduction_rmb",
        "ai_opex_reduction_pct",
        "annual_service_fee_rmb",
    ]
    ascending = [False, False, True, False, True, True, True]
    return df.sort_values(sort_cols, ascending=ascending, kind="mergesort").reset_index(drop=True)


def provenance_bundle(baseline: BaselineResults, best: LaaSResults | None) -> dict:
    out = {
        "baseline": {
            "revenue": baseline.revenue_rmb_y.provenance.as_dict(),
            "cash_opex": baseline.cash_opex_rmb_y.provenance.as_dict(),
            "electricity_opex": baseline.electricity_opex_rmb_y.provenance.as_dict(),
            "depreciation": baseline.depreciation_rmb_y.provenance.as_dict(),
            "capex": baseline.capex_provenance.as_dict(),
            "debt_service": baseline.debt_service_rmb_y.provenance.as_dict(),
            "gross_profit": baseline.accounting_gross_profit_rmb_y.provenance.as_dict(),
        }
    }
    if best is not None:
        out["laas_best"] = {
            "scenario": asdict(best.scenario),
            "client_payment": best.client_payment_rmb_y.provenance.as_dict(),
            "provider_cash_opex": best.provider_cash_opex_rmb_y.provenance.as_dict(),
            "provider_gross_profit": best.provider_accounting_gross_profit_rmb_y.provenance.as_dict(),
            "client_gap_rmb": {
                "value": float(best.client_gap_rmb),
                "units": "RMB (PV)",
                "definition": "client_gap = PV(LaaS payments incl upfront) − PV(baseline payments) − PV(ValueFromGuarantees). <=0 means client better off.",
                "components": {
                    "baseline_client_npv_cost_rmb": float(best.baseline_client_npv_cost_rmb),
                    "laas_client_npv_cost_rmb": float(best.laas_client_npv_cost_rmb),
                    "guarantees_npv_value_rmb": float(best.guarantees_npv_value_rmb),
                },
            },
            "feasibility": best.feasibility_provenance.as_dict(),
        }
    return out

