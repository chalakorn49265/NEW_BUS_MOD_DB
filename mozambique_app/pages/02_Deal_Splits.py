from __future__ import annotations

import sys
from pathlib import Path

import pandas as pd
import plotly.express as px
import streamlit as st

_ROOT = Path(__file__).resolve().parents[2]
if str(_ROOT) not in sys.path:
    sys.path.insert(0, str(_ROOT))

from mozambique.baseline import DEFAULT_BASELINES  # noqa: E402
from mozambique.deal import DealCharge, DealSplits, DealSubscription  # noqa: E402
from mozambique.scenario import MozambiqueInputs, run_mozambique_scenario  # noqa: E402


NAVY = "#1F3864"
MUTED = "#6B7280"
ACCENT = "#2563EB"


def _money(x: float | int | None) -> str:
    if x is None:
        return "-"
    try:
        return f"{float(x):,.0f}"
    except Exception:
        return "-"


def main() -> None:
    st.set_page_config(page_title="Mozambique — Deal Splits", layout="wide", initial_sidebar_state="expanded")
    st.markdown(
        f"<h2 style='color:{NAVY};margin-bottom:0.2rem;'>Deal splits (government, intermediaries, undertable)</h2>"
        f"<p style='color:{MUTED};margin-top:0;'>These flows are modeled as provider outflows and shown transparently.</p>",
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
        st.subheader("Subscription (client pays)")
        term = st.number_input("Term (years)", min_value=1, max_value=20, value=10, step=1)
        annual_fee = st.number_input("Annual subscription fee (USD/year)", min_value=0.0, value=600_000.0, step=25_000.0)
        upfront = st.number_input("Upfront payment (USD, month 0)", min_value=0.0, value=0.0, step=25_000.0)
        esc = st.slider("Annual escalation (%)", -5.0, 20.0, 3.0, 0.5) / 100.0

        st.divider()
        st.subheader("Provider economics")
        capex_total = st.number_input("Provider CAPEX total (USD)", min_value=0.0, value=3_000_000.0, step=100_000.0)
        opex_annual = st.number_input("Provider annual OPEX (USD/year)", min_value=0.0, value=120_000.0, step=10_000.0)

        st.divider()
        st.subheader("Split inputs")
        st.caption("Percentages apply to subscription inflow. Fixed amounts are annual unless marked upfront.")
        gov_pct = st.slider("Government (% of subscription)", 0.0, 40.0, 5.0, 0.5) / 100.0
        partner_pct = st.slider("Local partner (% of subscription)", 0.0, 40.0, 8.0, 0.5) / 100.0
        agent_pct = st.slider("Agent/introducer (% of subscription)", 0.0, 40.0, 0.0, 0.5) / 100.0
        undertable_pct = st.slider("Undertable (% of subscription)", 0.0, 40.0, 0.0, 0.5) / 100.0

        undertable_fixed = st.number_input("Undertable fixed (USD/year)", min_value=0.0, value=0.0, step=10_000.0)
        undertable_upfront = st.number_input("Undertable upfront (USD, month 0)", min_value=0.0, value=0.0, step=25_000.0)

    baseline = DEFAULT_BASELINES[baseline_type]  # type: ignore[index]
    subscription = DealSubscription(term_years=int(term), annual_fee_usd=float(annual_fee), upfront_usd=float(upfront), escalation_pct_annual=float(esc))
    charges = [
        DealCharge(recipient="Government", kind="government", timing="annual", pct_of_subscription=float(gov_pct)),
        DealCharge(recipient="Local partner", kind="intermediary", timing="annual", pct_of_subscription=float(partner_pct)),
        DealCharge(recipient="Agent/introducer", kind="intermediary", timing="annual", pct_of_subscription=float(agent_pct)),
    ]
    if undertable_pct > 0:
        charges.append(DealCharge(recipient="Undertable", kind="undertable", timing="annual", pct_of_subscription=float(undertable_pct)))
    if undertable_fixed > 0:
        charges.append(DealCharge(recipient="Undertable (fixed)", kind="undertable", timing="annual", fixed_usd=float(undertable_fixed)))
    if undertable_upfront > 0:
        charges.append(DealCharge(recipient="Undertable (upfront)", kind="undertable", timing="upfront", fixed_usd=float(undertable_upfront)))
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

    left, right = st.columns([1.2, 1.0])

    with left:
        st.subheader("Stakeholder revenue list (annual)")
        df = pd.DataFrame(res.stakeholder_revenue_annual)
        if df.empty:
            st.info("No splits configured.")
        else:
            df = df.sort_values(["year", "kind", "stakeholder"]).reset_index(drop=True)
            st.dataframe(df, use_container_width=True, height=420)

            piv = df.pivot_table(index=["year"], columns=["kind"], values="cash_in_usd", aggfunc="sum").fillna(0.0)
            piv["total"] = piv.sum(axis=1)
            fig = px.bar(
                piv.reset_index(),
                x="year",
                y=[c for c in piv.columns if c != "year"],
                title="Where the money goes (by year)",
                template="plotly_white",
            )
            fig.update_layout(height=360, font_color=NAVY, xaxis_title="Year", yaxis_title="USD/year")
            st.plotly_chart(fig, use_container_width=True)

    with right:
        st.subheader("Feasibility impact (provider net)")
        k1, k2, k3 = st.columns(3)
        k1.metric("IRR (net, annual)", str(res.provider_net_irr_annual) if isinstance(res.provider_net_irr_annual, str) else f"{res.provider_net_irr_annual:.1%}")
        k2.metric("Payback (months)", str(res.provider_net_payback_month))
        k3.metric("End-of-term cumulative", _money(res.provider_net_cumulative_annual[-1] if res.provider_net_cumulative_annual else 0.0))

        df_net = pd.DataFrame(
            {
                "year": list(range(0, len(res.provider_net_annual_y0_to_yN))),
                "net_cashflow": res.provider_net_annual_y0_to_yN,
                "cumulative": res.provider_net_cumulative_annual,
            }
        )
        st.dataframe(df_net, use_container_width=True, height=260)
        st.caption("Convention: Year 0 includes CAPEX outflow and upfront inflow and any upfront charges.")

        st.subheader("Client headline (sanity check)")
        baseline_total = res.baseline_energy_annual_usd + res.baseline_maintenance_annual_usd
        st.markdown(
            f"- **Baseline annual cost**: USD {_money(baseline_total)}\n"
            f"- **Subscription annual fee**: USD {_money(subscription.annual_fee_usd)}\n"
            f"- **Baseline type**: {baseline.name}"
        )


if __name__ == "__main__":
    main()

