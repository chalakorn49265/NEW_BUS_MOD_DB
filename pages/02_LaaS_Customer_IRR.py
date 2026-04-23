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
    CustomerBaselineInputs,
    CustomerLaaSInputs,
    SolveError,
    customer_incremental_cashflows_monthly,
    irr_annual_from_monthly_cashflows,
    payback_month_from_monthly_cashflows,
    solve_customer_for_target_irr,
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
            name="Cumulative net benefit",
            mode="lines",
            line=dict(color=ACCENT, width=3),
        )
    )
    fig.add_trace(
        go.Scatter(
            x=df["month"],
            y=df["unrecovered"],
            name="Unrecovered (benefit gap)",
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
        title="Cumulative net benefit vs baseline (Customer perspective)",
        font_color=NAVY,
        xaxis_title="Month",
        yaxis_title="USD",
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="left", x=0),
    )
    return fig


def main() -> None:
    st.set_page_config(page_title="LaaS — Customer IRR", layout="wide", initial_sidebar_state="expanded")
    st.markdown(
        f"<h2 style='color:{NAVY};margin-bottom:0.2rem;'>LaaS IRR Target — Customer</h2>"
        f"<p style='color:{MUTED};margin-top:0;'>Customer IRR is computed on incremental cashflows: "
        f"(baseline cost) − (LaaS payments + residual costs).</p>",
        unsafe_allow_html=True,
    )

    with st.sidebar:
        st.subheader("Target")
        target_irr = st.slider("Target IRR (annual)", -0.20, 0.80, 0.18, 0.005)
        solve_for = st.selectbox(
            "Solve-for parameter",
            ["annual_fee_usd", "upfront_usd", "term_years"],
            index=0,
            help="Solver adjusts ONLY this one knob to hit the target customer IRR.",
        )

        st.divider()
        st.subheader("Baseline (existing system)")
        term_years = st.number_input("Term (years)", min_value=1, max_value=30, value=10, step=1)
        b_energy = st.number_input("Baseline energy cost (USD/year)", min_value=0.0, value=900_000.0, step=50_000.0)
        b_maint = st.number_input(
            "Baseline maintenance cost (USD/year)", min_value=0.0, value=150_000.0, step=25_000.0
        )
        b_esc = st.slider("Baseline escalation (annual %)", -0.05, 0.20, 0.03, 0.005)

        st.divider()
        st.subheader("LaaS (new system)")
        annual_fee = st.number_input("Annual LaaS fee (USD/year)", min_value=0.0, value=600_000.0, step=50_000.0)
        upfront = st.number_input("Upfront payment (USD, paid at month 0)", value=0.0, step=50_000.0)
        esc = st.slider("LaaS fee escalation (annual %)", -0.05, 0.20, 0.03, 0.005)

        st.divider()
        st.subheader("Residual customer costs (post-upgrade)")
        r_energy = st.number_input("Residual energy cost (USD/year)", min_value=0.0, value=0.0, step=25_000.0)
        r_maint = st.number_input("Residual maintenance cost (USD/year)", min_value=0.0, value=0.0, step=25_000.0)
        r_esc = st.slider("Residual escalation (annual %)", -0.05, 0.20, 0.03, 0.005)

        st.divider()
        st.subheader("Solver bounds")
        b_lo = st.number_input("Lower bound (for solve-for, USD)", value=0.0, step=50_000.0)
        b_hi = st.number_input("Upper bound (for solve-for, USD)", value=5_000_000.0, step=250_000.0)

    baseline = CustomerBaselineInputs(
        term_years=int(term_years),
        baseline_energy_annual_usd=float(b_energy),
        baseline_maintenance_annual_usd=float(b_maint),
        baseline_escalation_pct_annual=float(b_esc),
    )
    laas = CustomerLaaSInputs(
        term_years=int(term_years),
        annual_fee_usd=float(annual_fee),
        upfront_usd=float(upfront),
        escalation_pct_annual=float(esc),
        residual_energy_annual_usd=float(r_energy),
        residual_maintenance_annual_usd=float(r_maint),
        residual_escalation_pct_annual=float(r_esc),
    )

    solved = None
    npv_at_target = None
    error: str | None = None
    try:
        solved, npv_at_target = solve_customer_for_target_irr(
            baseline,
            laas,
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

    flows = customer_incremental_cashflows_monthly(baseline, solved)
    irr_ach = irr_annual_from_monthly_cashflows(flows)
    payback_m = payback_month_from_monthly_cashflows(flows)

    months = list(range(0, len(flows)))
    net = pd.Series(flows, name="net_benefit")
    cum = net.cumsum()
    df = pd.DataFrame({"month": months, "net_benefit": net, "cumulative_net": cum})
    df["unrecovered"] = (-df["cumulative_net"]).clip(lower=0.0)

    k1, k2, k3, k4 = st.columns(4)
    irr_disp = irr_ach if isinstance(irr_ach, str) else f"{irr_ach:.1%}"
    k1.metric("Achieved IRR (annual)", str(irr_disp))
    k2.metric("NPV @ target IRR (USD)", f"{npv_at_target:,.0f}")
    k3.metric("Payback (months)", str(payback_m))
    if solve_for == "term_years":
        k4.metric("Solved term (years)", str(solved.term_years))
    elif solve_for == "annual_fee_usd":
        k4.metric("Solved annual fee (USD/yr)", f"{solved.annual_fee_usd:,.0f}")
    else:
        k4.metric("Solved upfront (USD)", f"{solved.upfront_usd:,.0f}")

    c1, c2 = st.columns([1.3, 1.0])
    with c1:
        st.plotly_chart(_cum_fig(df, payback_m), use_container_width=True)
    with c2:
        fig = px.bar(df.iloc[1:].head(36), x="month", y="net_benefit", title="Monthly net benefit (first 36 months)")
        fig.update_layout(template="plotly_white", height=430, font_color=NAVY, xaxis_title="Month", yaxis_title="USD")
        st.plotly_chart(fig, use_container_width=True)

    st.caption(
        "Convention: month 0 contains any upfront payment (negative to customer). Payback is the first month where cumulative net benefit ≥ 0."
    )


if __name__ == "__main__":
    main()

