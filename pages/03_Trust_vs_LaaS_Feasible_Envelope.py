from __future__ import annotations

import json
import sys
from pathlib import Path

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

_ROOT = Path(__file__).resolve().parents[1]
if str(_ROOT) not in sys.path:
    sys.path.insert(0, str(_ROOT))

from business_model_comparison.laas_feasible import (  # noqa: E402
    ClientValueAssumptions,
    MAX_AI_OPEX_REDUCTION_PCT,
    MIN_LAST_FOUR_YEAR_FEE_PCT,
    default_fee_grid_from_baseline,
    grid_search_feasible_envelope,
)
from business_model_comparison.models import build_baseline_energy_trust  # noqa: E402
from business_model_comparison.roadlight_data import load_roadlight_all  # noqa: E402
from business_model_comparison.report import (  # noqa: E402
    baseline_summary_table,
    envelope_table,
    laas_results_to_table,
    provenance_bundle,
    rank_recommended_offers,
)


NAVY = "#1F3864"
MUTED = "#6B7280"
ACCENT = "#2563EB"
GREEN = "#16A34A"
RED = "#DC2626"


def _money(x: float) -> str:
    return f"{x:,.0f}"


def main() -> None:
    st.set_page_config(page_title="能源托管 → LaaS (Feasible Envelope)", layout="wide", initial_sidebar_state="expanded")
    st.markdown(
        f"<h2 style='color:{NAVY};margin-bottom:0.2rem;'>能源托管 → LaaS : feasible offer envelope</h2>"
        f"<p style='color:{MUTED};margin-top:0;'>This page does not let users freely tweak outcomes. "
        "It enumerates a range of LaaS offers and shows only those that satisfy: term≤10y, provider payback≤36m, "
        "client pays less (each year), provider accounting gross profit higher (each year), and provider payback improves vs baseline. "
        "It can also evaluate a stepped handover-style fee where years 7-10 are reduced, subject to a non-zero floor. Every number has traceability.</p>",
        unsafe_allow_html=True,
    )

    with st.sidebar:
        st.subheader("Horizon & constraints")
        horizon_years = st.slider("Max term (years)", 1, 10, 10, 1)
        payback_m = st.slider("Provider payback constraint (months)", 6, 60, 36, 1)
        st.caption("Payback is simple cash payback on provider project cashflows (month 0 includes CAPEX and upfront).")

        st.divider()
        st.subheader("Offer search ranges")
        fee_low = st.slider("Service fee low (% of baseline 托管费)", 0.1, 1.0, 0.5, 0.05)
        fee_high = st.slider("Service fee high (% of baseline 托管费)", 0.2, 1.2, 1.0, 0.05)
        fee_steps = st.number_input("Fee grid steps", 10, 200, 60, 5)
        tail_reduction_grid = st.multiselect(
            "Last 4 years fee reduction grid (RMB/year)",
            options=[0.0, 200_000.0, 400_000.0, 600_000.0, 800_000.0, 1_000_000.0, 1_200_000.0, 1_500_000.0],
            default=[0.0, 400_000.0, 800_000.0],
            help=f"Applied only in years 7-10. The commercial fee in those years is floored at {int(MIN_LAST_FOUR_YEAR_FEE_PCT * 100)}% of the main annual fee so it cannot collapse to zero.",
        )

        upfront_grid = st.multiselect(
            "Upfront (RMB) grid",
            options=[0.0, 200_000.0, 500_000.0, 1_000_000.0, 2_000_000.0],
            default=[0.0],
            help="Upfront is paid by client at month 0 (positive to provider).",
        )
        st.divider()
        st.subheader("OPEX transformation")
        opex_modes = st.multiselect(
            "OPEX mode(s)",
            options=["uniform_pct", "electricity_only_pct", "ai_plus_solar"],
            default=["uniform_pct", "electricity_only_pct", "ai_plus_solar"],
            help="uniform_pct reduces total cash OPEX; electricity_only_pct reduces only 改造后电费; ai_plus_solar sets electricity OPEX to 0 (assumption; no solar CAPEX modeled).",
        )
        ai_grid = st.multiselect(
            "Reduction % grid (for pct modes)",
            options=[0, 5, 10, 15, 20, 25, 30, 35, 40, 60, 80, 85],
            default=[0, 10, 20, 30, 40, 60, 80, 85],
            help=f"Used only for modes that apply a percentage reduction. Percentage-style AI OPEX reduction is capped at {int(MAX_AI_OPEX_REDUCTION_PCT * 100)}%. For ai_plus_solar, electricity is forced to 0 regardless of this grid.",
        )

        st.divider()
        st.subheader("Valuation")
        disc = st.slider("Discount rate (annual, NPV)", 0.0, 0.30, 0.12, 0.005)

        st.divider()
        st.subheader("Client value (outcome guarantee)")
        b_out = st.number_input("Baseline outage hours/year", min_value=0.0, value=30.0, step=1.0)
        l_out = st.number_input("LaaS guaranteed outage hours/year", min_value=0.0, value=5.0, step=1.0)
        out_cost = st.number_input("Outage cost (RMB/hour)", min_value=0.0, value=10_000.0, step=1_000.0)
        share = st.slider("Credit/value share to client", 0.0, 1.0, 1.0, 0.05)
        client_disc = st.slider("Client discount rate (annual)", 0.0, 0.30, float(disc), 0.005)

        st.divider()
        run = st.button("Run search", type="primary")

    if not run:
        st.info("Set ranges in the sidebar, then click “Run search”.")
        return

    parsed = load_roadlight_all(_ROOT / "data")
    baseline = build_baseline_energy_trust(parsed, analysis_years=int(horizon_years), discount_rate_annual=float(disc))

    fee_grid = default_fee_grid_from_baseline(
        baseline,
        pct_low=float(fee_low),
        pct_high=float(fee_high),
        steps=int(fee_steps),
    )
    env = grid_search_feasible_envelope(
        baseline,
        term_years=list(range(1, int(horizon_years) + 1)),
        annual_fee_rmb_grid=fee_grid,
        last_four_year_fee_reduction_rmb_grid=[float(x) for x in tail_reduction_grid],
        upfront_rmb_grid=[float(x) for x in upfront_grid],
        ai_opex_reduction_grid=[float(x) / 100.0 for x in ai_grid],
        discount_rate_annual=float(disc),
        opex_modes=[x for x in opex_modes],  # type: ignore[arg-type]
        client_value=ClientValueAssumptions(
            baseline_outage_hours_per_year=float(b_out),
            laas_guaranteed_outage_hours_per_year=float(l_out),
            outage_cost_rmb_per_hour=float(out_cost),
            sla_credit_share_to_client=float(share),
            client_discount_rate_annual=float(client_disc),
        ),
    )

    # Enforce payback threshold (page-level). Core evaluator uses 36m; we filter for display.
    env_df = envelope_table(env)
    env_df["payback_ok_custom"] = env_df["payback_months"].apply(lambda x: (isinstance(x, int) and x <= int(payback_m)) or (isinstance(x, float) and x <= float(payback_m)))
    provider_df = env_df[(env_df["provider_feasible"]) & (env_df["payback_ok_custom"])]
    feasible_df = env_df[(env_df["feasible_everyone_better_off"]) & (env_df["payback_ok_custom"])]
    recommended_df = rank_recommended_offers(feasible_df)

    k1, k2, k3, k4 = st.columns(4)
    k1.metric("Evaluated scenarios", f"{env.evaluated:,}")
    k2.metric("Provider-feasible", f"{int(env_df['provider_feasible'].sum()):,}")
    k3.metric("Everyone-feasible (client+provider)", f"{int(env_df['feasible_everyone_better_off'].sum()):,}")
    k4.metric("Baseline payback (months)", str(baseline.payback_months))

    st.divider()
    tab_env, tab_best, tab_trace = st.tabs(["Feasible envelope", "Best feasible offer", "Traceability"])

    with tab_env:
        st.subheader("Offer map (term × fee) — show provider-feasible and client-gap")
        if provider_df.empty:
            st.warning("No provider-feasible offers under the current grid and constraints. Expand fee range, increase OPEX improvements, or allow upfront.")
        else:
            dfp = provider_df.copy()
            # Marker: everyone-feasible vs provider-only.
            dfp["layer"] = np.where(dfp["feasible_everyone_better_off"], "Everyone-feasible", "Provider-only")

            fig = px.scatter(
                dfp,
                x="annual_service_fee_rmb",
                y="term_years",
                color="client_gap_rmb",
                symbol="layer",
                facet_col="opex_mode",
                color_continuous_scale="RdYlGn_r",
                title="Provider-feasible offers (color = client_gap RMB; <=0 means client better off after guarantee value)",
                hover_data=[
                    "upfront_rmb",
                    "last_four_year_fee_reduction_rmb",
                    "ai_opex_reduction_pct",
                    "payback_months",
                    "npv_project_rmb",
                    "client_benefit_pass",
                    "meets_pay_less_each_year",
                ],
            )
            fig.update_layout(
                template="plotly_white",
                height=520,
                font_color=NAVY,
                xaxis_title="Annual service fee (RMB)",
                yaxis_title="Term (years)",
            )
            st.plotly_chart(fig, use_container_width=True)

            st.caption(
                "client_gap_rmb = PV(LaaS payments incl upfront) − PV(baseline payments) − PV(ValueFromGuarantees). "
                "Negative is good for client. Recommendation selection is stricter: client must also pay less each year, provider GP must be higher each year, and payback must improve vs baseline. "
                "When a last-4-year reduction is used, years 7-10 are stepped down but never to zero."
            )

            st.dataframe(
                dfp.sort_values(
                    ["term_years", "annual_service_fee_rmb", "last_four_year_fee_reduction_rmb", "opex_mode", "ai_opex_reduction_pct", "upfront_rmb"]
                ),
                use_container_width=True,
            )

    with tab_best:
        st.subheader("Best feasible offer (maximize client savings + 华普毛空间 uplift, then faster payback)")
        best_row = None if recommended_df.empty else recommended_df.head(1).iloc[0].to_dict()
        if best_row is None:
            st.info("No feasible offers to select from.")
        else:
            st.markdown(
                f"- **Term**: {int(best_row['term_years'])} years\n"
                f"- **Annual service fee**: RMB {_money(float(best_row['annual_service_fee_rmb']))}\n"
                f"- **Last 4 years reduction**: RMB {_money(float(best_row['last_four_year_fee_reduction_rmb']))} / year\n"
                f"- **Upfront**: RMB {_money(float(best_row['upfront_rmb']))}\n"
                f"- **OPEX mode**: {best_row['opex_mode']}\n"
                f"- **Reduction (pct modes)**: {float(best_row['ai_opex_reduction_pct']):.0%}\n"
                f"- **Payback**: {best_row['payback_months']} months\n"
                f"- **Payback improvement vs baseline**: {float(best_row['payback_improvement_months']):,.0f} months\n"
                f"- **Provider NPV**: RMB {_money(float(best_row['npv_project_rmb']))}\n"
                f"- **Client gap (RMB PV)**: {_money(float(best_row['client_gap_rmb']))} (<=0 means client better off)\n"
                f"- **Min client savings / year**: RMB {_money(float(best_row['min_client_savings_rmb_per_year']))}\n"
                f"- **Min 华普毛空间 uplift / year**: RMB {_money(float(best_row['min_provider_gross_profit_uplift_rmb_per_year']))}\n"
            )

            # Re-evaluate to get per-year detail tables.
            from business_model_comparison.laas_feasible import LaaSScenario, evaluate_laas_scenario

            best = evaluate_laas_scenario(
                baseline,
                LaaSScenario(
                    term_years=int(best_row["term_years"]),
                    annual_service_fee_rmb=float(best_row["annual_service_fee_rmb"]),
                    last_four_year_fee_reduction_rmb=float(best_row["last_four_year_fee_reduction_rmb"]),
                    upfront_rmb=float(best_row["upfront_rmb"]),
                    ai_opex_reduction_pct=float(best_row["ai_opex_reduction_pct"]),
                    opex_mode=str(best_row["opex_mode"]),  # type: ignore[arg-type]
                ),
                discount_rate_annual=float(disc),
                client_value=ClientValueAssumptions(
                    baseline_outage_hours_per_year=float(b_out),
                    laas_guaranteed_outage_hours_per_year=float(l_out),
                    outage_cost_rmb_per_hour=float(out_cost),
                    sla_credit_share_to_client=float(share),
                    client_discount_rate_annual=float(client_disc),
                ),
            )

            detail = laas_results_to_table(best, baseline)
            detail_operating = detail[detail["year"] > 0].copy()
            c1, c2 = st.columns(2)
            with c1:
                fig1 = go.Figure()
                fig1.add_trace(go.Bar(x=detail_operating["year"], y=detail_operating["baseline_trust_fee_rmb"], name="Baseline 托管费", marker_color="#94A3B8"))
                fig1.add_trace(go.Bar(x=detail_operating["year"], y=detail_operating["laas_service_fee_rmb"], name="LaaS 服务费", marker_color=ACCENT))
                fig1.update_layout(barmode="group", template="plotly_white", height=360, title="Client payment comparison (RMB/year)", font_color=NAVY)
                st.plotly_chart(fig1, use_container_width=True)
            with c2:
                fig2 = go.Figure()
                fig2.add_trace(go.Bar(x=detail_operating["year"], y=detail_operating["baseline_gross_profit_rmb"], name="Baseline 年毛空间", marker_color="#94A3B8"))
                fig2.add_trace(go.Bar(x=detail_operating["year"], y=detail_operating["laas_gross_profit_rmb"], name="LaaS 年毛空间", marker_color=GREEN))
                fig2.update_layout(barmode="group", template="plotly_white", height=360, title="Provider accounting gross profit (RMB/year)", font_color=NAVY)
                st.plotly_chart(fig2, use_container_width=True)

            st.dataframe(detail, use_container_width=True)

    with tab_trace:
        st.subheader("Traceability bundle (JSON)")
        # Show provenance for baseline, and best if exists.
        best = None
        if not recommended_df.empty:
            best_row = recommended_df.head(1).iloc[0].to_dict()
            from business_model_comparison.laas_feasible import LaaSScenario, evaluate_laas_scenario

            best = evaluate_laas_scenario(
                baseline,
                LaaSScenario(
                    term_years=int(best_row["term_years"]),
                    annual_service_fee_rmb=float(best_row["annual_service_fee_rmb"]),
                    last_four_year_fee_reduction_rmb=float(best_row["last_four_year_fee_reduction_rmb"]),
                    upfront_rmb=float(best_row["upfront_rmb"]),
                    ai_opex_reduction_pct=float(best_row["ai_opex_reduction_pct"]),
                ),
                discount_rate_annual=float(disc),
            )
        bundle = provenance_bundle(baseline, best)
        st.code(json.dumps(bundle, ensure_ascii=False, indent=2), language="json")

        st.subheader("Baseline yearly table (for audit)")
        st.dataframe(baseline_summary_table(baseline), use_container_width=True)


if __name__ == "__main__":
    main()

