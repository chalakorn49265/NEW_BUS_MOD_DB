from __future__ import annotations

import sys
from pathlib import Path

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

_ROOT = Path(__file__).resolve().parents[1]
if str(_ROOT) not in sys.path:
    sys.path.insert(0, str(_ROOT))

from emc_institutional_model.laas import (  # noqa: E402
    ProviderLaaSInputs,
    SolveError,
    irr_annual_from_monthly_cashflows,
    payback_month_from_monthly_cashflows,
    provider_cashflows_monthly,
    solve_provider_for_target_irr,
)

NAVY = "#1F3864"
MUTED = "#6B7280"
ACCENT = "#2563EB"
UNRECOVERED = "#EA580C"


def _cum_fig(df: pd.DataFrame, payback_m: int | str) -> go.Figure:
    fig = go.Figure()
    fig.add_trace(
        go.Scatter(
            x=df["month"],
            y=df["cumulative_net"],
            name="Cumulative net cashflow",
            mode="lines",
            line=dict(color=ACCENT, width=3),
        )
    )
    fig.add_trace(
        go.Scatter(
            x=df["month"],
            y=df["unrecovered_capex"],
            name="Unrecovered CAPEX",
            mode="lines",
            line=dict(color=UNRECOVERED, width=2, dash="dot"),
        )
    )
    if isinstance(payback_m, int):
        fig.add_vline(x=payback_m, line_width=2, line_dash="dash", line_color=MUTED)
        fig.add_annotation(x=payback_m, y=0, text="Payback", showarrow=True, ax=30, ay=-30, font=dict(color=NAVY))
    fig.update_layout(
        template="plotly_white",
        height=430,
        title="Cumulative net cashflow (Provider perspective)",
        font_color=NAVY,
        xaxis_title="Month",
        yaxis_title="USD",
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="left", x=0),
    )
    return fig


def main() -> None:
    st.set_page_config(page_title="LaaS — Provider IRR", layout="wide", initial_sidebar_state="expanded")
    st.markdown(
        f"<h2 style='color:{NAVY};margin-bottom:0.2rem;'>LaaS IRR Target — Provider / Project</h2>"
        f"<p style='color:{MUTED};margin-top:0;'>Set a target provider IRR and the model solves a selected parameter "
        f"(annual fee, upfront, or term). CAPEX is fixed at $3.0M and provider OPEX is fixed at $0 in this view.</p>",
        unsafe_allow_html=True,
    )

    with st.sidebar:
        st.subheader("Target")
        target_irr = st.slider("Target IRR (annual)", -0.20, 0.80, 0.18, 0.005)
        solve_for = st.selectbox(
            "Solve-for parameter",
            ["annual_fee_usd", "upfront_usd", "term_years"],
            index=0,
            help="To avoid infinite solutions, the solver adjusts ONLY this one knob to hit the target IRR.",
        )

        st.divider()
        st.subheader("Fixed baseline")
        st.number_input("CAPEX (fixed)", value=3_000_000.0, disabled=True, help="Locked per requirement.")

        st.divider()
        st.subheader("Contract & costs")
        term_years = st.number_input("Term (years)", min_value=1, max_value=30, value=10, step=1)
        annual_fee = st.number_input("Annual customer fee (USD/year)", min_value=0.0, value=600_000.0, step=50_000.0)
        upfront = st.number_input("Upfront payment (USD, received at month 0)", value=0.0, step=50_000.0)
        esc = st.slider("Escalation (annual %)", -0.05, 0.20, 0.03, 0.005)

        st.divider()
        st.subheader("Solver bounds")
        b_lo = st.number_input("Lower bound (for solve-for, USD)", value=0.0, step=50_000.0)
        b_hi = st.number_input("Upper bound (for solve-for, USD)", value=5_000_000.0, step=250_000.0)

    base = ProviderLaaSInputs(
        capex_usd=3_000_000.0,
        term_years=int(term_years),
        annual_fee_usd=float(annual_fee),
        upfront_usd=float(upfront),
        escalation_pct_annual=float(esc),
        provider_opex_annual_usd=0.0,
    )

    solved = None
    npv_at_target = None
    error: str | None = None

    try:
        solved, npv_at_target = solve_provider_for_target_irr(
            base,
            target_irr_annual=float(target_irr),
            solve_for=solve_for,  # type: ignore[arg-type]
            bounds=(float(b_lo), float(b_hi)),
            term_year_bounds=(1, 30),
        )
    except SolveError as e:
        error = str(e)

    if error or solved is None or npv_at_target is None:
        st.error(f"Cannot solve under current constraints. {error or ''}".strip())
        return

    flows = provider_cashflows_monthly(solved)
    irr_ach = irr_annual_from_monthly_cashflows(flows)
    payback_m = payback_month_from_monthly_cashflows(flows)

    months = list(range(0, len(flows)))
    net = pd.Series(flows, name="net_cashflow")
    cum = net.cumsum()
    df = pd.DataFrame({"month": months, "net_cashflow": net, "cumulative_net": cum})
    df["unrecovered_capex"] = (-df["cumulative_net"]).clip(lower=0.0)

    k1, k2, k3, k4 = st.columns(4)
    irr_disp = irr_ach if isinstance(irr_ach, str) else f"{irr_ach:.1%}"
    k1.metric("Achieved IRR (annual)", str(irr_disp))
    k2.metric("NPV @ target IRR (USD)", f"{npv_at_target:,.0f}")
    k3.metric("Payback (months)", str(payback_m))
    if solve_for == "term_years":
        k4.metric("Solved term (years)", str(solved.term_years))
    elif solve_for == "annual_fee_usd":
        k4.metric("Solved annual fee (USD/yr)", f"{solved.annual_fee_usd:,.0f}")
    elif solve_for == "upfront_usd":
        k4.metric("Solved upfront (USD)", f"{solved.upfront_usd:,.0f}")
    else:
        k4.metric("Solved parameter", str(solve_for))

    c1, c2 = st.columns([1.3, 1.0])
    with c1:
        st.plotly_chart(_cum_fig(df, payback_m), use_container_width=True)
    with c2:
        fig = px.bar(df.iloc[1:].head(36), x="month", y="net_cashflow", title="Monthly net cashflow (first 36 months)")
        fig.update_layout(template="plotly_white", height=430, font_color=NAVY, xaxis_title="Month", yaxis_title="USD")
        st.plotly_chart(fig, use_container_width=True)

    st.caption(
        "Convention: month 0 contains CAPEX outflow and any upfront payment. Payback is the first month where cumulative net ≥ 0."
    )


if __name__ == "__main__":
    main()

