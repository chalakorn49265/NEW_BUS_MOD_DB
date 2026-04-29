from __future__ import annotations

import json
import sys
from pathlib import Path

import pandas as pd
import streamlit as st

_ROOT = Path(__file__).resolve().parents[2]
if str(_ROOT) not in sys.path:
    sys.path.insert(0, str(_ROOT))

from mozambique.baseline import DEFAULT_BASELINES  # noqa: E402
from mozambique.deal import DealCharge, DealSplits, DealSubscription  # noqa: E402
from mozambique.distribution import annualize_monthly  # noqa: E402
from mozambique.scenario import MozambiqueInputs, run_mozambique_scenario  # noqa: E402


NAVY = "#1F3864"
MUTED = "#6B7280"


def main() -> None:
    st.set_page_config(page_title="Mozambique — Audit", layout="wide", initial_sidebar_state="expanded")
    st.markdown(
        f"<h2 style='color:{NAVY};margin-bottom:0.2rem;'>Audit view (inputs → outputs)</h2>"
        f"<p style='color:{MUTED};margin-top:0;'>Detailed tables so the sales team can answer questions rigorously.</p>",
        unsafe_allow_html=True,
    )

    with st.sidebar:
        st.subheader("Project")
        baseline_type = st.selectbox("Existing baseline type", options=["LED", "HPS"], index=1)
        n = st.number_input("Number of lights", min_value=1, value=1000, step=50)
        hours = st.slider("Operating hours per night", 4.0, 14.0, 11.0, 0.25)
        days = st.number_input("Days per year", min_value=300.0, max_value=366.0, value=365.0, step=1.0)
        price = st.number_input("Electricity price (USD/kWh)", min_value=0.0, value=0.10, step=0.01, format="%.2f")

        st.divider()
        st.subheader("Offer (subscription)")
        term = st.number_input("Term (years)", min_value=1, max_value=20, value=10, step=1)
        annual_fee = st.number_input("Annual subscription fee (USD/year)", min_value=0.0, value=600_000.0, step=25_000.0)
        upfront = st.number_input("Upfront payment (USD, month 0)", min_value=0.0, value=0.0, step=25_000.0)
        esc = st.slider("Annual escalation (%)", -5.0, 20.0, 3.0, 0.5) / 100.0

        st.divider()
        st.subheader("Provider economics")
        capex_total = st.number_input("Provider CAPEX total (USD)", min_value=0.0, value=3_000_000.0, step=100_000.0)
        opex_annual = st.number_input("Provider annual OPEX (USD/year)", min_value=0.0, value=120_000.0, step=10_000.0)

        st.divider()
        st.subheader("Charges (simple)")
        gov_pct = st.slider("Government (% of subscription)", 0.0, 40.0, 5.0, 0.5) / 100.0
        interm_pct = st.slider("Intermediaries (% of subscription)", 0.0, 40.0, 8.0, 0.5) / 100.0
        undertable = st.number_input("Undertable fixed (USD/year)", min_value=0.0, value=0.0, step=10_000.0)

    baseline = DEFAULT_BASELINES[baseline_type]  # type: ignore[index]
    subscription = DealSubscription(term_years=int(term), annual_fee_usd=float(annual_fee), upfront_usd=float(upfront), escalation_pct_annual=float(esc))
    charges = [
        DealCharge(recipient="Government", kind="government", timing="annual", pct_of_subscription=float(gov_pct)),
        DealCharge(recipient="Intermediaries", kind="intermediary", timing="annual", pct_of_subscription=float(interm_pct)),
    ]
    if undertable > 0:
        charges.append(DealCharge(recipient="Undertable", kind="undertable", timing="annual", fixed_usd=float(undertable)))
    splits = DealSplits(charges=charges)

    res = run_mozambique_scenario(
        baseline=baseline,
        subscription=subscription,
        splits=splits,
        inputs=MozambiqueInputs(
            number_of_lights=int(n),
            operating_hours_per_night=float(hours),
            days_per_year=float(days),
            electricity_price_usd_per_kwh=float(price),
            capex_usd_total=float(capex_total),
            provider_opex_annual_usd=float(opex_annual),
        ),
    )

    tab_inputs, tab_customer, tab_provider, tab_splits, tab_trace = st.tabs(
        ["Inputs", "Customer view", "Provider view", "Splits", "Traceability"]
    )

    with tab_inputs:
        st.subheader("Inputs snapshot")
        st.json(
            {
                "baseline": baseline.__dict__,
                "subscription": subscription.__dict__,
                "provider": {"capex_usd_total": capex_total, "provider_opex_annual_usd": opex_annual},
                "project": {
                    "number_of_lights": n,
                    "operating_hours_per_night": hours,
                    "days_per_year": days,
                    "electricity_price_usd_per_kwh": price,
                },
                "charges": [c.__dict__ for c in splits.normalized().charges],
            }
        )

    with tab_customer:
        st.subheader("Customer incremental cashflows (baseline − LaaS)")
        annual = annualize_monthly(res.customer_incremental_monthly)
        df = pd.DataFrame({"year": list(range(0, len(annual))), "incremental_benefit_usd": annual})
        st.dataframe(df, use_container_width=True, height=360)
        st.caption("Positive values mean the client is better off vs baseline in that year.")

    with tab_provider:
        st.subheader("Provider net cashflows (after splits)")
        df = pd.DataFrame(
            {
                "year": list(range(0, len(res.provider_net_annual_y0_to_yN))),
                "net_cashflow_usd": res.provider_net_annual_y0_to_yN,
                "cumulative_usd": res.provider_net_cumulative_annual,
            }
        )
        st.dataframe(df, use_container_width=True, height=420)
        st.caption("Year 0 includes CAPEX and upfront and upfront charges.")

    with tab_splits:
        st.subheader("Stakeholder revenues (annual)")
        df = pd.DataFrame(res.stakeholder_revenue_annual)
        st.dataframe(df.sort_values(["year", "kind", "stakeholder"]), use_container_width=True, height=420)

    with tab_trace:
        st.subheader("Traceability (model conventions)")
        st.markdown(
            "- **Baseline**: annual cost = electricity + maintenance (explicit LED/HPS tech).\n"
            "- **AI+Solar**: client residual energy & maintenance are modeled as 0 in this dashboard.\n"
            "- **Provider**: cashflow = subscription inflow − provider OPEX − charges; month 0 includes CAPEX and upfront.\n"
            "- **Charges**: pct charges apply to subscription inflow; fixed charges are annual; all are provider outflows.\n"
        )
        st.subheader("Traceability bundle (JSON)")
        st.code(json.dumps(res.traceability, ensure_ascii=False, indent=2), language="json")


if __name__ == "__main__":
    main()

