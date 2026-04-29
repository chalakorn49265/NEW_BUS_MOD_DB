from __future__ import annotations

import sys
from pathlib import Path

import pandas as pd
import plotly.graph_objects as go
import streamlit as st

_ROOT = Path(__file__).resolve().parents[2]
if str(_ROOT) not in sys.path:
    sys.path.insert(0, str(_ROOT))

from mozambique.baseline import DEFAULT_BASELINES  # noqa: E402
from mozambique.deal import DealCharge, DealSplits, DealSubscription  # noqa: E402
from mozambique.evidence import default_evidence_cards  # noqa: E402
from mozambique.scenario import MozambiqueInputs, run_mozambique_scenario  # noqa: E402


NAVY = "#1F3864"
MUTED = "#6B7280"
ACCENT = "#2563EB"
GREEN = "#16A34A"
RED = "#DC2626"


def _money(x: float | int | None) -> str:
    if x is None:
        return "-"
    try:
        return f"{float(x):,.0f}"
    except Exception:
        return "-"


def _pct(x: float | str | None) -> str:
    if x is None:
        return "-"
    if isinstance(x, str):
        return x
    try:
        return f"{float(x):.1%}"
    except Exception:
        return "-"


def main() -> None:
    st.set_page_config(page_title="Mozambique — 1-min Pitch", layout="wide", initial_sidebar_state="expanded")
    st.markdown(
        f"<h2 style='color:{NAVY};margin-bottom:0.2rem;'>AI+Solar Lighting (LaaS) — 1 minute pitch</h2>"
        f"<p style='color:{MUTED};margin-top:0;'>Before: client pays electricity + maintenance. "
        f"After: client pays one subscription fee, electricity is ~0 (solar) and provider handles maintenance.</p>",
        unsafe_allow_html=True,
    )

    with st.sidebar:
        st.subheader("Project (simple inputs)")
        baseline_type = st.selectbox("Existing baseline type", options=["LED", "HPS"], index=1)
        n = st.number_input("Number of lights", min_value=1, value=1000, step=50)
        hours = st.slider("Operating hours per night", 4.0, 14.0, 11.0, 0.25)
        days = st.number_input("Days per year", min_value=300.0, max_value=366.0, value=365.0, step=1.0)
        price = st.number_input("Electricity price (USD/kWh)", min_value=0.0, value=0.10, step=0.01, format="%.2f")

        st.divider()
        st.subheader("Offer (client subscription)")
        term = st.number_input("Term (years)", min_value=1, max_value=20, value=10, step=1)
        annual_fee = st.number_input("Annual subscription fee (USD/year)", min_value=0.0, value=600_000.0, step=25_000.0)
        upfront = st.number_input("Upfront payment (USD, month 0)", min_value=0.0, value=0.0, step=25_000.0)
        esc = st.slider("Annual escalation (%)", -5.0, 20.0, 3.0, 0.5) / 100.0

        st.divider()
        st.subheader("Provider economics (for feasibility)")
        capex_total = st.number_input("Provider CAPEX total (USD)", min_value=0.0, value=3_000_000.0, step=100_000.0)
        opex_annual = st.number_input(
            "Provider annual OPEX (USD/year) — maintenance, operations",
            min_value=0.0,
            value=120_000.0,
            step=10_000.0,
        )

        st.divider()
        st.subheader("Quick splits (optional)")
        gov_pct = st.slider("Government share (% of subscription)", 0.0, 30.0, 5.0, 0.5) / 100.0
        interm_pct = st.slider("Intermediaries (% of subscription)", 0.0, 30.0, 8.0, 0.5) / 100.0
        undertable_annual = st.number_input("Undertable (USD/year)", min_value=0.0, value=0.0, step=10_000.0)

    baseline = DEFAULT_BASELINES[baseline_type]  # type: ignore[index]
    subscription = DealSubscription(
        term_years=int(term),
        annual_fee_usd=float(annual_fee),
        upfront_usd=float(upfront),
        escalation_pct_annual=float(esc),
    )
    charges = [
        DealCharge(recipient="Government", kind="government", timing="annual", pct_of_subscription=float(gov_pct)),
        DealCharge(recipient="Intermediaries", kind="intermediary", timing="annual", pct_of_subscription=float(interm_pct)),
    ]
    if undertable_annual > 0:
        charges.append(
            DealCharge(recipient="Undertable", kind="undertable", timing="annual", fixed_usd=float(undertable_annual))
        )
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

    # Block A: Customer before/after
    st.subheader("A) Client (Before → After)")
    c1, c2, c3, c4 = st.columns(4)
    baseline_total = res.baseline_energy_annual_usd + res.baseline_maintenance_annual_usd
    c1.metric("Before: electricity (USD/yr)", _money(res.baseline_energy_annual_usd))
    c2.metric("Before: maintenance (USD/yr)", _money(res.baseline_maintenance_annual_usd))
    c3.metric("Before: total (USD/yr)", _money(baseline_total))
    c4.metric("After: subscription (USD/yr)", _money(subscription.annual_fee_usd))

    st.caption(
        f"Baseline selected: **{baseline.name}**. Under AI+Solar LaaS we assume **grid electricity ~0** and provider handles maintenance; the client pays the subscription."
    )

    # Block B: Where subscription money goes
    st.subheader("B) Where the subscription money goes (transparent)")
    df_st = pd.DataFrame(res.stakeholder_revenue_annual)
    if df_st.empty:
        st.info("No splits configured.")
    else:
        piv = df_st.pivot_table(index=["year"], columns=["kind", "stakeholder"], values="cash_in_usd", aggfunc="sum").fillna(0.0)
        st.dataframe(piv, use_container_width=True)

    # Block C: Provider feasibility
    st.subheader("C) Provider feasibility (subscription must cover CAPEX + OPEX + splits)")
    p1, p2, p3, p4 = st.columns(4)
    p1.metric("Provider IRR (net, annual)", _pct(res.provider_net_irr_annual))
    p2.metric("Provider payback (months)", str(res.provider_net_payback_month))
    y10 = res.provider_net_cumulative_annual[-1] if res.provider_net_cumulative_annual else 0.0
    p3.metric("Cumulative net (end of term)", _money(y10))
    p4.metric("Upfront (month 0)", _money(subscription.upfront_usd))

    years = list(range(0, len(res.provider_net_cumulative_annual)))
    fig = go.Figure()
    fig.add_trace(go.Scatter(x=years, y=res.provider_net_cumulative_annual, mode="lines+markers", name="Provider cumulative net", line=dict(color=ACCENT, width=3)))
    fig.add_hline(y=0, line_dash="dash", line_color="#CBD5E1")
    fig.update_layout(template="plotly_white", height=360, font_color=NAVY, xaxis_title="Year (Y0..)", yaxis_title="USD")
    st.plotly_chart(fig, use_container_width=True)

    with st.expander("Assumptions (show if asked)"):
        st.markdown(
            "- **Before**: baseline annual cost = electricity + maintenance (explicit baseline tech).\n"
            "- **After**: AI+Solar LaaS → client pays subscription; grid electricity is treated as ~0; maintenance is on provider OPEX.\n"
            "- **Splits**: government / intermediary / undertable are modeled as provider outflows (reducing provider IRR), shown transparently."
        )

    with st.expander("Why trust the numbers (evidence)"):
        for c in default_evidence_cards():
            with st.expander(c.title, expanded=False):
                st.markdown(f"**Why**: {c.why}")
                st.markdown(f"**Evidence**: {c.evidence}")
                if c.notes:
                    st.markdown(f"**Notes**: {c.notes}")


if __name__ == "__main__":
    main()

