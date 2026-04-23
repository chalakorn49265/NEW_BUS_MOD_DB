"""Streamlit sales dashboard for the EMC institutional model."""

from __future__ import annotations

import sys
from typing import Union
from pathlib import Path

_ROOT = Path(__file__).resolve().parent
if str(_ROOT) not in sys.path:
    sys.path.insert(0, str(_ROOT))

from datetime import date

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import streamlit as st

from emc_institutional_model.cashflow import MonthRow, payback_month, payback_month_adjusted
from emc_institutional_model.defaults import PRODUCT_ROWS
from emc_institutional_model.energy import implied_energy_savings_fraction
from emc_institutional_model.metrics import scenario_kpis
from emc_institutional_model.monte_carlo import MonteCarloConfig, fan_chart_quantiles, run_monte_carlo
from emc_institutional_model.params import EmcAdjustments, ModelParams, TariffModel, TouSide
from emc_institutional_model.runner import run_model

NAVY = "#1F3864"
MUTED = "#6B7280"
ACCENT = "#2563EB"
UNRECOVERED = "#EA580C"
RED_GAP = "rgba(239,68,68,0.15)"


def _tou_from_flat(gov: float, edm: float) -> TariffModel:
    return TariffModel.flat(gov, edm)


def _tou_street_default(gov_peak: float, gov_off: float, edm_peak: float, edm_off: float) -> TariffModel:
    """Night-heavy load: mostly off-peak for street lighting."""
    w = [0.75, 0.20, 0.05]
    gov_side = TouSide(
        bucket_ids=["offpeak", "shoulder", "peak"],
        prices_usd_per_kwh=[gov_off, (gov_off + gov_peak) / 2, gov_peak],
        load_weights=w,
    )
    edm_side = TouSide(
        bucket_ids=["offpeak", "shoulder", "peak"],
        prices_usd_per_kwh=[edm_off, (edm_off + edm_peak) / 2, edm_peak],
        load_weights=w,
    )
    return TariffModel(gov_payment=gov_side, edm_cost=edm_side)


def _scale_gov_prices(p: ModelParams, factor: float) -> ModelParams:
    g = p.tariff_model.gov_payment
    new_g = TouSide(
        bucket_ids=list(g.bucket_ids),
        prices_usd_per_kwh=[x * factor for x in g.prices_usd_per_kwh],
        load_weights=list(g.load_weights),
    )
    return p.model_copy(update={"tariff_model": TariffModel(gov_payment=new_g, edm_cost=p.tariff_model.edm_cost)})


def _scale_edm_prices(p: ModelParams, factor: float) -> ModelParams:
    e = p.tariff_model.edm_cost
    new_e = TouSide(
        bucket_ids=list(e.bucket_ids),
        prices_usd_per_kwh=[x * factor for x in e.prices_usd_per_kwh],
        load_weights=list(e.load_weights),
    )
    return p.model_copy(update={"tariff_model": TariffModel(gov_payment=p.tariff_model.gov_payment, edm_cost=new_e)})


def _tornado_npv(base: ModelParams, swing: float = 0.12) -> pd.DataFrame:
    b = run_model(base).npv_usd
    rows = []
    for label, up, down in [
        ("Gov tariff ×(1±s)", _scale_gov_prices(base, 1 + swing), _scale_gov_prices(base, 1 - swing)),
        ("EDM tariff ×(1±s)", _scale_edm_prices(base, 1 + swing), _scale_edm_prices(base, 1 - swing)),
        (
            "Hours ×(1±s)",
            base.model_copy(update={"operating_hours_per_night": base.operating_hours_per_night * (1 + swing)}),
            base.model_copy(update={"operating_hours_per_night": base.operating_hours_per_night * (1 - swing)}),
        ),
        (
            "Escalation +2pt",
            base.model_copy(update={"escalation_pct_annual": base.escalation_pct_annual + 0.02}),
            base.model_copy(update={"escalation_pct_annual": max(-0.05, base.escalation_pct_annual - 0.02)}),
        ),
        (
            "Horizon +24m",
            base.model_copy(update={"analysis_length_months": min(600, base.analysis_length_months + 24)}),
            base.model_copy(update={"analysis_length_months": max(12, base.analysis_length_months - 24)}),
        ),
    ]:
        nu = run_model(up).npv_usd
        nd = run_model(down).npv_usd
        rows.append(
            {
                "driver": label,
                "npv_up": nu,
                "npv_down": nd,
                "tornado_up": nu - b,
                "tornado_down": nd - b,
            }
        )
    return pd.DataFrame(rows)


def _emc_adjustments_active(fp: EmcAdjustments) -> bool:
    return (
        fp.corporate_tax_rate > 0
        or fp.distribution_pct_of_gross_inflow > 0
        or fp.distribution_fixed_usd_month > 0
    )


def _prepare_sources_uses_frame(monthly: pd.DataFrame, cadence: str) -> pd.DataFrame:
    if cadence == "Monthly":
        out = monthly.copy()
        out["period_label"] = out["month_index"].astype(str)
        out["period_x"] = out["month_index"].astype(int)
        return out
    d = monthly.copy()
    d["year"] = (d["month_index"] - 1) // 12 + 1
    sum_cols = [
        "revenue_energy",
        "revenue_custody",
        "revenue_performance_fee",
        "capex_outflow",
        "electrical_fee",
        "maintenance_fee",
        "tax_cash",
        "distribution_fees",
        "net_cashflow",
        "net_cashflow_adjusted",
    ]
    g = d.groupby("year", as_index=False)[sum_cols].sum()
    g["period_label"] = "Y" + g["year"].astype(str)
    g["period_x"] = g["year"].astype(int)
    end_rows = d.loc[d.groupby("year")["month_index"].idxmax()][
        ["year", "cumulative_net_cashflow", "cumulative_net_adjusted"]
    ].rename(
        columns={
            "cumulative_net_cashflow": "cumulative_net_cashflow_end",
            "cumulative_net_adjusted": "cumulative_net_adjusted_end",
        }
    )
    g = g.merge(end_rows, on="year", how="left")
    g["cumulative_net_cashflow"] = g["cumulative_net_cashflow_end"]
    g["cumulative_net_adjusted"] = g["cumulative_net_adjusted_end"]
    g = g.drop(columns=["cumulative_net_cashflow_end", "cumulative_net_adjusted_end"], errors="ignore")
    return g


def _sources_uses_figure(
    df: pd.DataFrame,
    cum_col: str,
    net_col: str,
    payback_period: Union[int, str],
    x_title: str,
    cumulative_subtitle: str,
) -> go.Figure:
    """Two-row layout: bars on top; cumulative net + unrecovered equipment CAPEX on bottom."""
    x = df["period_x"]
    unrecovered = (-df[cum_col].astype(float)).clip(lower=0.0)
    fig = make_subplots(
        rows=2,
        cols=1,
        shared_xaxes=True,
        vertical_spacing=0.12,
        row_heights=[0.55, 0.45],
        subplot_titles=("Sources & uses (period cash)", cumulative_subtitle),
        specs=[[{"secondary_y": False}], [{"secondary_y": False}]],
    )
    pos = [
        ("Energy revenue", "revenue_energy", "#16A34A"),
        ("Custody", "revenue_custody", "#4ADE80"),
        ("Performance fee", "revenue_performance_fee", "#86EFAC"),
    ]
    neg = [
        ("Electrical OPEX", "electrical_fee", "#F97316"),
        ("Maintenance / M&V", "maintenance_fee", "#FB923C"),
        ("Taxes", "tax_cash", "#DC2626"),
        ("Distribution / fees", "distribution_fees", "#7C3AED"),
        ("CAPEX (equipment)", "capex_outflow", "#1E3A5F"),
    ]
    for name, col, c in pos:
        fig.add_trace(
            go.Bar(x=x, y=df[col], name=name, marker_color=c, legendgroup="src", showlegend=True),
            row=1,
            col=1,
        )
    for name, col, c in neg:
        fig.add_trace(
            go.Bar(x=x, y=df[col], name=name, marker_color=c, legendgroup="use", showlegend=True),
            row=1,
            col=1,
        )

    fig.add_trace(
        go.Scatter(
            x=x,
            y=df[cum_col],
            name="Cumulative net",
            mode="lines+markers",
            line=dict(color=ACCENT, width=2),
            marker=dict(size=5),
            showlegend=True,
        ),
        row=2,
        col=1,
    )
    fig.add_trace(
        go.Scatter(
            x=x,
            y=unrecovered,
            name="Unrecovered equipment CAPEX",
            mode="lines+markers",
            line=dict(color=UNRECOVERED, width=2, dash="dot"),
            marker=dict(size=4),
            showlegend=True,
        ),
        row=2,
        col=1,
    )

    shapes = []
    for _, row in df.iterrows():
        if float(row[net_col]) < 0:
            pxv = float(row["period_x"])
            shapes.append(
                dict(
                    type="rect",
                    xref="x",
                    yref="y domain",
                    x0=pxv - 0.45,
                    x1=pxv + 0.45,
                    y0=0,
                    y1=1,
                    fillcolor=RED_GAP,
                    line_width=0,
                    layer="below",
                )
            )

    fig.update_layout(
        barmode="relative",
        template="plotly_white",
        font_color=NAVY,
        height=680,
        legend=dict(orientation="h", yanchor="bottom", y=1.06, xanchor="right", x=1),
        shapes=shapes,
    )
    fig.update_yaxes(title_text="USD (period)", row=1, col=1)
    # Single y-axis for both lines so y=0 is one horizontal line (orange hits 0 when blue crosses 0).
    fig.update_yaxes(
        title_text="Cumulative net & unrecovered equipment (USD)",
        row=2,
        col=1,
        zeroline=True,
        zerolinewidth=1.5,
        zerolinecolor="#64748B",
    )
    fig.update_xaxes(title_text=x_title, row=2, col=1)

    if isinstance(payback_period, int):
        fig.add_shape(
            type="line",
            x0=payback_period,
            x1=payback_period,
            y0=0,
            y1=1,
            xref="x",
            yref="paper",
            line=dict(color=MUTED, width=2, dash="dash"),
        )
        mask = df["period_x"] == payback_period
        y_pb = float(df.loc[mask, cum_col].iloc[0]) if mask.any() else 0.0
        fig.add_annotation(
            x=payback_period,
            y=y_pb,
            xref="x2",
            yref="y2",
            text="Payback",
            showarrow=True,
            arrowhead=2,
            ax=30,
            ay=-25,
            font=dict(size=11, color=NAVY),
        )
    return fig


def main() -> None:
    st.set_page_config(page_title="EMC Financial Model", layout="wide", initial_sidebar_state="expanded")
    st.markdown(
        f"<h1 style='color:{NAVY};margin-bottom:0.2rem;'>EMC street lighting — financial cockpit</h1>"
        f"<p style='color:{MUTED};margin-top:0;'>Decision-grade engine (workbook parity + TOU + Monte Carlo). "
        "Figures are model outputs, not legal advice.</p>",
        unsafe_allow_html=True,
    )

    with st.sidebar:
        st.header("Project")
        start = st.date_input("Start date", value=date(2026, 1, 1))
        months = st.number_input("Horizon (months)", 12, 600, 120, 1)
        lights = st.number_input("Number of lights", 1, 500_000, 1000, 1)
        poles = st.number_input("Number of poles", 1, 500_000, 1000, 1)
        hours = st.slider("Operating hours / night", 4.0, 14.0, 11.0, 0.25)
        kw = st.number_input("Power kW / light", 0.01, 2.0, 0.10, 0.01)
        tier = st.selectbox("Location area", ["city_center", "suburb", "rural"])
        product = st.selectbox("Product", list(PRODUCT_ROWS.keys()))

        if "_track_product" not in st.session_state:
            st.session_state._track_product = product
        if product != st.session_state._track_product:
            st.session_state._track_product = product
            st.session_state._savings_pct = implied_energy_savings_fraction(product) * 100.0
        if "_savings_pct" not in st.session_state:
            st.session_state._savings_pct = implied_energy_savings_fraction(product) * 100.0

        savings_pct = st.slider(
            "Energy savings % ≈ (baseline kWh − delivered kWh) / baseline kWh",
            0.0,
            99.0,
            value=float(st.session_state._savings_pct),
            step=0.1,
            help="Overrides delivered consumption vs product baseline; drives avoided kWh and EDM OPEX.",
        )
        st.session_state._savings_pct = savings_pct

        basis = st.selectbox("Revenue basis", ["avoided_kwh", "delivered_kwh"])
        esc = st.slider("Annual escalation", -0.05, 0.20, 0.03, 0.005)
        custody_en = st.checkbox("Enable custody (托管) fee", value=False)
        custody = st.number_input("Custody fee USD / pole / month", 0.0, 500.0, 0.0, 1.0)
        perf_pct = st.slider("EMC performance fee on energy savings (% trad−AI EDM)", 0.0, 50.0, 0.0, 1.0) / 100.0
        disc = st.slider("Discount rate (annual, NPV)", 0.04, 0.25, 0.12, 0.005)
        st.divider()
        with st.expander("Tax & fees (Sources & Uses chart, optional)", expanded=False):
            st.caption("NPV / IRR in metrics stay on full equipment CAPEX; optional tax and fee lines adjust the chart only.")
            tax_rate = st.slider("Corporate tax rate on energy + performance (ex-custody)", 0.0, 0.40, 0.0, 0.01)
            dep_mo = st.number_input("Depreciation months (tax)", 1, 600, 120, 1)
            dist_pct = st.slider("Distribution % of gross contract inflow", 0.0, 0.50, 0.0, 0.005)
            dist_fix = st.number_input("Distribution / fee USD / month", 0.0, 100_000.0, 0.0, 100.0)
        st.divider()
        st.subheader("Tariffs")
        mode = st.radio("Tariff mode", ["Flat (legacy B4/B5)", "TOU night-heavy"], horizontal=True)
        if mode.startswith("Flat"):
            gov_f = st.number_input("Gov payment USD/kWh", 0.0, 2.0, 0.18, 0.01)
            edm_f = st.number_input("EDM cost USD/kWh", 0.0, 2.0, 0.10, 0.01)
            tariff_model = _tou_from_flat(gov_f, edm_f)
        else:
            st.caption("Peak vs off-peak; load shape fixed 75% / 20% / 5% (street-lighting default).")
            gov_off = st.number_input("Gov off-peak USD/kWh", 0.0, 2.0, 0.12, 0.01)
            gov_pk = st.number_input("Gov peak USD/kWh", 0.0, 2.0, 0.28, 0.01)
            edm_off = st.number_input("EDM off-peak USD/kWh", 0.0, 2.0, 0.07, 0.01)
            edm_pk = st.number_input("EDM peak USD/kWh", 0.0, 2.0, 0.18, 0.01)
            tariff_model = _tou_street_default(gov_pk, gov_off, edm_pk, edm_off)
        st.divider()
        st.subheader("Monte Carlo")
        run_mc = st.checkbox("Run Monte Carlo", value=False)
        mc_n = st.number_input("Paths", 50, 5000, 300, 50)
        mc_seed = st.number_input("RNG seed", 0, 999999, 42, 1)

    emc_adjustments = EmcAdjustments(
        corporate_tax_rate=float(tax_rate),
        depreciation_months=int(dep_mo),
        distribution_pct_of_gross_inflow=float(dist_pct),
        distribution_fixed_usd_month=float(dist_fix),
    )

    p = ModelParams(
        project_start_date=start,
        analysis_length_months=int(months),
        number_of_lights=int(lights),
        number_of_poles=int(poles),
        operating_hours_per_night=float(hours),
        power_kw_per_light=float(kw),
        location_tier=tier,  # type: ignore[arg-type]
        product_type=product,
        revenue_basis=basis,  # type: ignore[arg-type]
        escalation_pct_annual=float(esc),
        custody_fee_usd_per_pole_month=float(custody),
        custody_fee_enabled=custody_en,
        tariff_model=tariff_model,
        discount_rate_annual=float(disc),
        emc_performance_fee_pct_of_energy_savings=float(perf_pct),
        energy_savings_fraction=float(savings_pct / 100.0),
        emc_adjustments=emc_adjustments,
    )

    res = run_model(p)
    k = scenario_kpis(p)

    k1, k2, k3, k4 = st.columns(4)
    k1.metric("NPV (USD)", f"{res.npv_usd:,.0f}")
    irr_disp = res.irr_annual if isinstance(res.irr_annual, float) else str(res.irr_annual)
    k2.metric("IRR (annual)", f"{irr_disp}" if isinstance(irr_disp, str) else f"{irr_disp:.1%}")
    k3.metric("Payback (months)", str(k["est_payback_months"]))
    k4.metric("Required custody fee ($/pole/mo)", f"{k['required_custody_fee_usd_per_pole_month']:.2f}")

    tab_cf, tab_wf, tab_tor, tab_tco, tab_opex, tab_tou, tab_mc = st.tabs(
        ["Cashflow", "Sources & Uses", "Tornado", "TCO bridge", "OPEX", "TOU", "Monte Carlo"]
    )

    with tab_cf:
        fig = px.area(
            res.monthly,
            x="month_index",
            y="cumulative_net_cashflow",
            title="Cumulative net cashflow (USD) — equipment CAPEX basis",
        )
        fig.update_layout(template="plotly_white", font_color=NAVY, height=420)
        st.plotly_chart(fig, use_container_width=True)
        if run_mc:
            mc = run_monte_carlo(p, MonteCarloConfig(n_paths=int(mc_n), seed=int(mc_seed)))
            fan = fan_chart_quantiles(mc.monthly_net_panel)
            m = res.monthly["month_index"]
            fig2 = go.Figure()
            fig2.add_trace(
                go.Scatter(x=m, y=res.monthly["net_cashflow"], name="Base case net", line=dict(color=ACCENT))
            )
            fig2.add_trace(
                go.Scatter(
                    x=m,
                    y=fan["p95"],
                    mode="lines",
                    line=dict(width=0),
                    showlegend=False,
                    hoverinfo="skip",
                )
            )
            fig2.add_trace(
                go.Scatter(
                    x=m,
                    y=fan["p5"],
                    fill="tonexty",
                    fillcolor="rgba(37,99,235,0.12)",
                    mode="lines",
                    name="MC monthly net P5–P95",
                )
            )
            fig2.update_layout(template="plotly_white", title="Monthly net cashflow vs MC band", height=420)
            st.plotly_chart(fig2, use_container_width=True)

    with tab_wf:
        cadence = st.radio("Period", ["Monthly", "Annual"], horizontal=True)
        adj_active = _emc_adjustments_active(emc_adjustments)
        cum_col = "cumulative_net_adjusted" if adj_active else "cumulative_net_cashflow"
        net_col = "net_cashflow_adjusted" if adj_active else "net_cashflow"
        df_plot = _prepare_sources_uses_frame(res.monthly, cadence)
        rows_list = res.monthly.to_dict("records")
        mr = [MonthRow(**{kk: float(r[kk]) for kk in MonthRow.__dataclass_fields__}) for r in rows_list]
        pb_m = payback_month_adjusted(mr) if adj_active else payback_month(mr)
        if isinstance(pb_m, int):
            pb_x: Union[int, str] = pb_m if cadence == "Monthly" else (pb_m - 1) // 12 + 1
        else:
            pb_x = "NO_PAYBACK"
        x_title = "Month index" if cadence == "Monthly" else "Year (from start)"
        cum_sub = (
            "Cumulative — after tax & fees (payback matches this line)"
            if adj_active
            else "Cumulative — equipment project cash (same basis as Payback months KPI)"
        )
        figw = _sources_uses_figure(df_plot, cum_col, net_col, pb_x, x_title, cum_sub)
        figw.update_layout(title=dict(text="Sources & Uses + cumulative net (EMC)", font=dict(color=NAVY)))
        st.plotly_chart(figw, use_container_width=True)
        if isinstance(pb_m, str):
            st.caption("No payback in horizon under the selected cumulative definition.")
        else:
            pb_disp = int(pb_m) if isinstance(pb_m, (int, float)) else pb_m
            st.caption(
                f"Payback: first month with cumulative net at or above zero is month {pb_disp}. "
                "Shaded columns: periods with negative net cashflow."
                if cadence == "Monthly"
                else (
                    f"Payback: first month at or above zero is month {pb_disp} "
                    f"(year {(int(pb_m) - 1) // 12 + 1} on this annual chart). "
                    "Shaded columns: negative net periods."
                )
            )
        st.caption(
            "Blue: cumulative net (USD). Orange: unrecovered equipment CAPEX = max(0, -cumulative net). "
            "Both use the same vertical axis, so the horizontal zero line is shared: orange reaches zero in the "
            "same month cumulative net first reaches zero. Revenue and OPEX are already in cumulative net."
        )
        if adj_active:
            st.caption("Cumulative basis includes tax and fee adjustments from the sidebar expander.")
        else:
            st.caption("Cumulative basis matches the Payback months KPI at the top of the page.")

    with tab_tor:
        tor = _tornado_npv(p)
        figt = go.Figure()
        figt.add_bar(y=tor["driver"], x=tor["tornado_up"], name="Upside swing", marker_color=ACCENT)
        figt.add_bar(y=tor["driver"], x=tor["tornado_down"], name="Downside swing", marker_color="#94A3B8")
        figt.update_layout(
            barmode="overlay",
            template="plotly_white",
            title="Tornado — ΔNPV vs base (± swings)",
            xaxis_title="Δ NPV (USD)",
            height=400,
        )
        st.plotly_chart(figt, use_container_width=True)

    with tab_tco:
        ai_capex = float(k["est_capex_usd"])
        tr_capex = float(k["traditional_capex_usd"])
        ai_opex_y = float(k["est_monthly_opex_usd"]) * 12
        tr_opex_y = float(k["traditional_monthly_opex_usd"]) * 12
        ai_rev_y = float(k["est_monthly_revenue_usd"]) * 12
        tr_rev_y = float(k["traditional_monthly_revenue_usd"]) * 12
        bridge = pd.DataFrame(
            {
                "label": ["AI CAPEX", "Traditional CAPEX", "AI net OPEX (yr1)", "Trad net OPEX (yr1)", "AI revenue (yr1)", "Trad revenue (yr1)"],
                "AI": [ai_capex, 0, ai_opex_y, 0, -ai_rev_y, 0],
                "Traditional": [0, tr_capex, 0, tr_opex_y, 0, -tr_rev_y],
            }
        )
        figb = go.Figure()
        figb.add_trace(go.Bar(name="AI", x=bridge["label"], y=bridge["AI"], marker_color=NAVY))
        figb.add_trace(go.Bar(name="Traditional", x=bridge["label"], y=bridge["Traditional"], marker_color="#94A3B8"))
        figb.update_layout(barmode="group", template="plotly_white", title="Annualised cost/revenue snapshot (yr1 style)", height=440)
        st.plotly_chart(figb, use_container_width=True)
        st.caption("Negative revenue bars represent inflows (cash convention).")

    with tab_opex:
        om = res.monthly[["month_index", "electrical_fee", "maintenance_fee"]].head(min(60, len(res.monthly)))
        figo = go.Figure()
        figo.add_trace(go.Bar(x=om["month_index"], y=-om["electrical_fee"], name="Electrical (outflow)"))
        figo.add_trace(go.Bar(x=om["month_index"], y=-om["maintenance_fee"], name="Maintenance (outflow)"))
        figo.update_layout(barmode="stack", template="plotly_white", title="OPEX composition (first 60 months)", height=400)
        st.plotly_chart(figo, use_container_width=True)

    with tab_tou:
        g = p.tariff_model.gov_payment
        dfw = pd.DataFrame(
            {"bucket": g.bucket_ids, "load_weight": g.load_weights, "gov_USD_per_kWh": g.prices_usd_per_kwh}
        )
        figq = px.bar(dfw, x="bucket", y="load_weight", title="Load allocation by TOU bucket", color_discrete_sequence=[ACCENT])
        figq.update_layout(template="plotly_white", height=360)
        st.plotly_chart(figq, use_container_width=True)
        figp = px.line(dfw, x="bucket", y="gov_USD_per_kWh", markers=True, title="Contract price by bucket")
        figp.update_traces(line_color=NAVY)
        figp.update_layout(template="plotly_white", height=320)
        st.plotly_chart(figp, use_container_width=True)

    with tab_mc:
        if not run_mc:
            st.info("Enable “Run Monte Carlo” in the sidebar.")
        else:
            mc = run_monte_carlo(p, MonteCarloConfig(n_paths=int(mc_n), seed=int(mc_seed)))
            c1, c2 = st.columns(2)
            figh1 = px.histogram(mc.npv_samples, nbins=40, title="NPV distribution", color_discrete_sequence=[NAVY])
            figh1.update_layout(template="plotly_white")
            c1.plotly_chart(figh1, use_container_width=True)
            irr_ok = mc.irr_samples[np.isfinite(mc.irr_samples)]
            figh2 = px.histogram(irr_ok, nbins=40, title="IRR distribution (annual)", color_discrete_sequence=[ACCENT])
            figh2.update_layout(template="plotly_white")
            c2.plotly_chart(figh2, use_container_width=True)
            st.dataframe(mc.summary(), use_container_width=True)


if __name__ == "__main__":
    main()
