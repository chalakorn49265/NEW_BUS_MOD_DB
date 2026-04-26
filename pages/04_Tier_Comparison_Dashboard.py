from __future__ import annotations

import json
import sys
from pathlib import Path

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

_ROOT = Path(__file__).resolve().parents[1]
if str(_ROOT) not in sys.path:
    sys.path.insert(0, str(_ROOT))

from Dashboard_LaaS_vs_NYTG.tier_dashboard_data import (  # noqa: E402
    build_tier_tables,
    tier_traceability_dict,
)
from Dashboard_LaaS_vs_NYTG.evidence_cn import cards_for_selected_tier  # noqa: E402


NAVY = "#1F3864"
MUTED = "#6B7280"
ACCENT = "#2563EB"
GREEN = "#16A34A"
RED = "#DC2626"

# Reduce metric font size so large numbers fit.
st.markdown(
    """
<style>
  div[data-testid="metric-container"] * { line-height: 1.05; }
  div[data-testid="stMetricValue"] { font-size: 1.35rem; }
  div[data-testid="stMetricLabel"] { font-size: 0.9rem; }
</style>
""",
    unsafe_allow_html=True,
)


def _money(x: float | None) -> str:
    if x is None or x != x:
        return "-"
    return f"{x:,.0f}"


def _pct(x: float | None) -> str:
    if x is None or x != x:
        return "-"
    return f"{x:.2%}"


@st.cache_data(show_spinner=False)
def _load(new_models_dir: str, cache_bust: str) -> tuple[pd.DataFrame, pd.DataFrame, list[dict]]:
    df_sum, df_long, tiers = build_tier_tables(new_models_dir=Path(new_models_dir))
    tiers_dict = [tier_traceability_dict(t) for t in tiers]
    return df_sum, df_long, tiers_dict


def main() -> None:
    st.set_page_config(page_title="EMC → LaaS | 单方案查看", layout="wide", initial_sidebar_state="expanded")
    st.markdown(
        f"<h2 style='color:{NAVY};margin-bottom:0.2rem;'>能源托管/EMC → LaaS：单方案“工作簿查看器”（中文）</h2>"
        f"<p style='color:{MUTED};margin-top:0;'>选择一个方案后，本页直接读取该工作簿的缓存数值（与WPS打开看到一致），并以更直观的方式展示："
        f"<code>01_Dashboard</code> 核心KPI、<code>05_Annual_Model</code> 年度现金流/节省、以及 OPEX/CAPEX 结构与“为什么更好”的证据说明。</p>",
        unsafe_allow_html=True,
    )

    default_dir = str(_ROOT / "Dashboard_LaaS_vs_NYTG" / "new_models")
    with st.sidebar:
        st.subheader("数据源（来自 new_models）")
        new_models_dir = st.text_input("new_models folder", value=default_dir)
        view_mode = st.radio("展示模式", options=["只看关键结论（1分钟版）", "展开明细（投委会版）"], index=0)
        show_all_years_table = st.checkbox("显示年度表格（像Excel）", value=(view_mode != "只看关键结论（1分钟版）"))

    df_sum, df_long, tiers_dict = _load(new_models_dir, cache_bust="v3_big_table_provider_lines")

    if df_sum.empty:
        st.warning("No workbooks found. Check the folder path.")
        return

    # Single-tier selector
    tier_list = df_sum["tier"].tolist()
    sel = st.selectbox("选择方案（切换工作簿）", options=tier_list, index=0)

    dsel = df_sum[df_sum["tier"] == sel].iloc[0].to_dict()
    df_t = df_long[df_long["tier"] == sel].copy()

    # 1-minute: KPI cards
    st.subheader("A) 关键结论（1分钟版）")
    k1, k2, k3, k4 = st.columns(4)
    k1.metric("服务商 NPV（EMC → LaaS）", f"{_money(dsel.get('emc_npv'))} → {_money(dsel.get('laas_npv'))}")
    k2.metric("服务商 IRR（EMC → LaaS）", f"{_pct(dsel.get('emc_irr'))} → {_pct(dsel.get('laas_irr'))}")
    k3.metric("业主净节省 Y1（EMC → LaaS）", f"{_money(dsel.get('owner_save_emc_y1'))} → {_money(dsel.get('owner_save_laas_y1'))}")
    k4.metric("业主10年累计净节省（EMC → LaaS）", f"{_money(dsel.get('owner_save_emc_sum10'))} → {_money(dsel.get('owner_save_laas_sum10'))}")

    st.caption(f"工作簿路径：`{dsel['file_path']}`（数据来自该工作簿缓存值，便于审计与对齐WPS）")

    st.divider()

    st.subheader("对比总表（像Excel，便于汇报）")
    # A wide table similar to the screenshot: year 0..10, EMC vs LaaS core cashflow building blocks.
    years = [0] + list(df_t["year"].tolist())
    trust_capex = [-(dsel.get("capex_emc_per_lamp") or 0.0) * float(dsel.get("lamps") or 0.0)] + [0.0] * 10
    laas_capex = [-(dsel.get("capex_laas_per_lamp") or 0.0) * float(dsel.get("lamps") or 0.0)] + [0.0] * 10
    # Owner total spend rows already mirror Annual_Model; use 0 for year0.
    trust_spend = [0.0] + [float(x or 0.0) for x in df_t["owner_spend_emc"].tolist()]
    laas_spend = [0.0] + [float(x or 0.0) for x in df_t["owner_spend_laas"].tolist()]
    trust_save = [0.0] + [float(x or 0.0) for x in df_t["owner_save_emc"].tolist()]
    laas_save = [0.0] + [float(x or 0.0) for x in df_t["owner_save_laas"].tolist()]
    laas_fee_sched = [0.0] + [float(x or 0.0) for x in df_t["laas_fee"].tolist()]
    # Use fee inputs as a proxy for EMC annual fee for display (workbook-only; actual annual fee schedule is flat in template).
    emc_fee = float(dsel.get("emc_fee_y1") or 0.0)
    trust_fee = [0.0] + [emc_fee] * 10
    laas_fee = [0.0] + laas_fee_sched[1:]

    def _compute_provider_block(*, kind: str) -> dict[str, list[float]]:
        """
        Build a provider-side cashflow block (Y0..Y10) using workbook inputs and
        the annual fee schedule already extracted from 05_Annual_Model.
        This avoids relying on cached values in formula-heavy rows that may be blank.
        """
        lamps = float(dsel.get("lamps") or 0.0)
        capex_per_lamp = float(dsel.get("capex_laas_per_lamp") or 0.0) if kind == "laas" else float(dsel.get("capex_emc_per_lamp") or 0.0)
        capex_y0 = capex_per_lamp * lamps

        # baseline electricity (Y1) - compute same as cost tab fallback
        baseline_elec_y1 = dsel.get("baseline_electricity_y1")
        if baseline_elec_y1 is None or baseline_elec_y1 != baseline_elec_y1:
            price = float(dsel.get("electricity_price_per_kwh") or 0.0)
            watts = float(dsel.get("watts_per_lamp") or 0.0)
            hpd = float(dsel.get("hours_per_day") or 0.0)
            dpy = float(dsel.get("days_per_year") or 0.0)
            baseline_elec_y1 = lamps * watts / 1000.0 * hpd * dpy * price if (lamps and watts and hpd and dpy and price) else 0.0
        baseline_elec_y1 = float(baseline_elec_y1 or 0.0)

        saving = float(dsel.get("laas_saving_rate") or 0.0) if kind == "laas" else float(dsel.get("emc_saving_rate") or 0.0)
        elec_after = [0.0] + [baseline_elec_y1 * (1.0 - saving)] * 10

        # Provider bears electricity cost only when switch is 1 in template; EMC typically 0.
        owner_pays_flag = float(dsel.get("emc_owner_pays_elec_flag") or 0.0)
        provider_bears_elec = (owner_pays_flag == 1.0) if kind == "emc" else True  # LaaS assumed provider bears unless workbook says otherwise
        elec_cost = [0.0] + ([elec_after[i] for i in range(1, 11)] if provider_bears_elec else [0.0] * 10)

        # OPEX components (non-electric)
        om = float(dsel.get("opex_om_per_lamp") or 0.0) * lamps
        platform = float(dsel.get("opex_platform") or 0.0)
        spares = float(dsel.get("opex_spares") or 0.0)
        om_y = [0.0] + [om] * 10
        platform_y = [0.0] + [platform] * 10
        spares_y = [0.0] + [spares] * 10

        # Revenue: service fee
        if kind == "laas":
            fee = [0.0] + [float(x or 0.0) for x in df_t["laas_fee"].tolist()]
        else:
            emc_fee = float(dsel.get("emc_fee_y1") or 0.0)
            fee = [0.0] + [emc_fee] * 10

        other_income = [0.0] * 11
        asset_income = [0.0] * 11

        total_rev = [fee[i] + other_income[i] + asset_income[i] for i in range(11)]
        total_cost = [elec_cost[i] + om_y[i] + platform_y[i] + spares_y[i] for i in range(11)]
        net_cf = [0.0] * 11
        net_cf[0] = -capex_y0
        for i in range(1, 11):
            net_cf[i] = total_rev[i] - total_cost[i]
        cum = []
        s = 0.0
        for i in range(11):
            s += net_cf[i]
            cum.append(s)

        return {
            "初始CAPEX支出（正数，扣减）": [capex_y0] + [0.0] * 10,
            "固定服务费收入": fee,
            "其他/第三方收入": other_income,
            "资产转让收入": asset_income,
            "节电后电费总额": elec_after,
            "服务商承担电费成本（正数，扣减）": elec_cost,
            "运维成本（正数，扣减）": om_y,
            "平台/管理成本（正数，扣减）": platform_y,
            "备件/小改造储备（正数，扣减）": spares_y,
            "年度总收入": total_rev,
            "年度总现金成本": total_cost,
            "年度净现金流": net_cf,
            "累计现金流": cum,
        }

    table_dict = {
        "year": years,
        # Owner view (top part)
        "EMC_capex(owner_view)": trust_capex,
        "EMC_fee(owner_view)": trust_fee,
        "EMC_total_spend(owner)": trust_spend,
        "EMC_net_savings(owner)": trust_save,
        "LaaS_capex(owner_view)": laas_capex,
        "LaaS_fee(owner_view)": laas_fee,
        "LaaS_total_spend(owner)": laas_spend,
        "LaaS_net_savings(owner)": laas_save,
    }
    prov_emc = _compute_provider_block(kind="emc")
    prov_laas = _compute_provider_block(kind="laas")
    for lab in prov_emc.keys():
        table_dict[f"EMC_服务商_{lab}"] = prov_emc[lab]
        table_dict[f"LaaS_服务商_{lab}"] = prov_laas[lab]

    table = pd.DataFrame(table_dict)

    st.caption("建议：默认仅展示关键列（适合汇报）；需要完整明细再展开。")
    show_full = st.toggle("展开：显示全部列（明细/审计）", value=False)

    key_cols = [
        "year",
        # Owner headline
        "EMC_total_spend(owner)",
        "LaaS_total_spend(owner)",
        "EMC_net_savings(owner)",
        "LaaS_net_savings(owner)",
        # Provider headline
        "EMC_服务商_年度净现金流",
        "LaaS_服务商_年度净现金流",
        "EMC_服务商_累计现金流",
        "LaaS_服务商_累计现金流",
    ]
    key_cols = [c for c in key_cols if c in table.columns]
    t_show = table if show_full else table[key_cols].copy()

    # Add deltas in key view
    if not show_full:
        if {"EMC_total_spend(owner)", "LaaS_total_spend(owner)"} <= set(t_show.columns):
            t_show.insert(
                3,
                "Δ业主总支出(LaaS-EMC)",
                t_show["LaaS_total_spend(owner)"] - t_show["EMC_total_spend(owner)"],
            )
        if {"EMC_net_savings(owner)", "LaaS_net_savings(owner)"} <= set(t_show.columns):
            t_show.insert(
                6,
                "Δ业主净节省(LaaS-EMC)",
                t_show["LaaS_net_savings(owner)"] - t_show["EMC_net_savings(owner)"],
            )
        if {"EMC_服务商_年度净现金流", "LaaS_服务商_年度净现金流"} <= set(t_show.columns):
            t_show["Δ服务商净现金流(LaaS-EMC)"] = t_show["LaaS_服务商_年度净现金流"] - t_show["EMC_服务商_年度净现金流"]

    def _fmt_money(v: float) -> str:
        try:
            return f"{float(v):,.0f}"
        except Exception:
            return "-"

    def _style(df: pd.DataFrame) -> "pd.io.formats.style.Styler":
        sty = df.style
        # Base number formatting
        num_cols = [c for c in df.columns if c != "year"]
        sty = sty.format({c: _fmt_money for c in num_cols})

        # Color-code columns so it's obvious which model is which
        emc_cols = [c for c in df.columns if c.startswith("EMC_")]
        laas_cols = [c for c in df.columns if c.startswith("LaaS_")]
        delta_cols = [c for c in df.columns if c.startswith("Δ")]

        if emc_cols:
            sty = sty.set_properties(
                subset=emc_cols,
                **{
                    "background-color": "#F8FAFC",  # very light slate
                    "color": "#334155",  # slate-700
                },
            )
        if laas_cols:
            sty = sty.set_properties(
                subset=laas_cols,
                **{
                    "background-color": "#ECFDF5",  # very light green
                    "color": "#065F46",  # green-800
                    "font-weight": "600",
                },
            )
        # Add a clear divider before the first LaaS column
        if laas_cols:
            first_laas = laas_cols[0]
            sty = sty.set_properties(subset=[first_laas], **{"border-left": "3px solid #10B981"})
        # Add a clear divider before delta block if present
        if delta_cols:
            first_delta = delta_cols[0]
            sty = sty.set_properties(subset=[first_delta], **{"border-left": "3px solid #F59E0B"})

        # Highlight most important columns
        highlight_cols = [c for c in df.columns if ("净节省" in c or "累计现金流" in c or "Δ" in c)]
        if highlight_cols:
            sty = sty.set_properties(subset=highlight_cols, **{"background-color": "#FFF7ED"})  # light orange

        # Conditional colors for deltas / savings
        def _color_pos_neg(val: float) -> str:
            try:
                x = float(val)
            except Exception:
                return ""
            if x > 0:
                return "color: #16A34A; font-weight: 600;"
            if x < 0:
                return "color: #DC2626; font-weight: 600;"
            return "color: #0F172A;"

        if delta_cols:
            sty = sty.map(_color_pos_neg, subset=delta_cols)

        save_cols = [c for c in df.columns if c.endswith("net_savings(owner)") or "净节省" in c]
        if save_cols:
            sty = sty.map(_color_pos_neg, subset=save_cols)

        return sty

    st.dataframe(_style(t_show), use_container_width=True, height=420)

    # Cumulative cashflow chart (provider) to show payback timing clearly
    st.subheader("累计现金流（服务商视角）：LaaS vs EMC（看回本更直观）")
    cum_emc = prov_emc.get("累计现金流", [0.0] * 11)
    cum_laas = prov_laas.get("累计现金流", [0.0] * 11)
    df_cum = pd.DataFrame({"year": years, "EMC": cum_emc, "LaaS": cum_laas})
    fig_cum = go.Figure()
    fig_cum.add_trace(go.Scatter(x=df_cum["year"], y=df_cum["EMC"], name="EMC 累计现金流", mode="lines+markers", line=dict(color="#94A3B8", width=3)))
    fig_cum.add_trace(go.Scatter(x=df_cum["year"], y=df_cum["LaaS"], name="LaaS 累计现金流", mode="lines+markers", line=dict(color=GREEN, width=3)))
    fig_cum.add_hline(y=0, line_dash="dash", line_color="#CBD5E1")
    fig_cum.update_layout(template="plotly_white", height=360, font_color=NAVY, xaxis_title="年（Y0..Y10）", yaxis_title="累计现金流（元）")
    st.plotly_chart(fig_cum, use_container_width=True)

    # OPEX factor story chart (Y1) - why LaaS is smaller
    st.subheader("为什么LaaS的OPEX更小：拆解对比（Y1，电费 + 非电费运维）")
    lamps = float(dsel.get("lamps") or 0.0)
    om_y1 = float(dsel.get("opex_om_per_lamp") or 0.0) * lamps
    platform_y1 = float(dsel.get("opex_platform") or 0.0)
    spares_y1 = float(dsel.get("opex_spares") or 0.0)

    baseline_elec_y1 = dsel.get("baseline_electricity_y1")
    if baseline_elec_y1 is None or baseline_elec_y1 != baseline_elec_y1:
        price = float(dsel.get("electricity_price_per_kwh") or 0.0)
        watts = float(dsel.get("watts_per_lamp") or 0.0)
        hpd = float(dsel.get("hours_per_day") or 0.0)
        dpy = float(dsel.get("days_per_year") or 0.0)
        baseline_elec_y1 = lamps * watts / 1000.0 * hpd * dpy * price if (lamps and watts and hpd and dpy and price) else 0.0
    baseline_elec_y1 = float(baseline_elec_y1 or 0.0)

    emc_save = float(dsel.get("emc_saving_rate") or 0.0)
    laas_save = float(dsel.get("laas_saving_rate") or 0.0)
    elec_emc_y1 = baseline_elec_y1 * (1.0 - emc_save)
    elec_laas_y1 = baseline_elec_y1 * (1.0 - laas_save)

    # Baseline: no savings (electricity stays baseline), non-electric OPEX shown as the same template components for storytelling consistency.
    df_opex_story = pd.DataFrame(
        [
            {"方案": "Baseline(现状)", "电费": baseline_elec_y1, "人工/维修": om_y1, "平台": platform_y1, "备件": spares_y1},
            {"方案": "EMC(托管)", "电费": elec_emc_y1, "人工/维修": om_y1, "平台": platform_y1, "备件": spares_y1},
            {"方案": "LaaS", "电费": elec_laas_y1, "人工/维修": om_y1, "平台": platform_y1, "备件": spares_y1},
        ]
    )
    df_m = df_opex_story.melt(id_vars=["方案"], var_name="成本项", value_name="金额")
    fig_opex = px.bar(
        df_m,
        x="方案",
        y="金额",
        color="成本项",
        barmode="stack",
        template="plotly_white",
        height=420,
        title="OPEX拆解对比（Y1）",
        color_discrete_map={"电费": "#60A5FA", "人工/维修": "#94A3B8", "平台": "#FBBF24", "备件": "#A78BFA"},
    )
    fig_opex.update_layout(font_color=NAVY, yaxis_title="元/年")
    st.plotly_chart(fig_opex, use_container_width=True)
    st.caption("注：电费来自“基准电费 × (1-节电率)”。非电费运维拆解来自模板输入（按灯人工/维修、平台费、备件）。如需“现状非电费运维”单独口径，我们可以从 `03_Baseline` 再补充提取。")

    tab_owner, tab_provider, tab_cost, tab_story, tab_trace = st.tabs(
        ["业主视角（年度表）", "服务商视角（回报）", "成本拆解（OPEX/CAPEX）", "为什么更好（证据）", "可追溯性"],
    )

    with tab_owner:
        st.subheader("B) 业主视角：年度总支出与年度净节省（对应 05_Annual_Model）")
        fig1 = go.Figure()
        fig1.add_trace(go.Scatter(x=df_t["year"], y=df_t["owner_spend_emc"], name="EMC 业主年度总支出", mode="lines+markers", line=dict(color="#94A3B8", width=3)))
        fig1.add_trace(go.Scatter(x=df_t["year"], y=df_t["owner_spend_laas"], name="LaaS 业主年度总支出", mode="lines+markers", line=dict(color=ACCENT, width=3)))
        fig1.update_layout(template="plotly_white", height=360, font_color=NAVY, xaxis_title="年", yaxis_title="元/年")
        st.plotly_chart(fig1, use_container_width=True)

        fig2 = go.Figure()
        fig2.add_trace(go.Scatter(x=df_t["year"], y=df_t["owner_save_emc"], name="EMC 业主年度净节省", mode="lines+markers", line=dict(color=RED, width=3)))
        fig2.add_trace(go.Scatter(x=df_t["year"], y=df_t["owner_save_laas"], name="LaaS 业主年度净节省", mode="lines+markers", line=dict(color=GREEN, width=3)))
        fig2.update_layout(template="plotly_white", height=360, font_color=NAVY, xaxis_title="年", yaxis_title="元/年")
        st.plotly_chart(fig2, use_container_width=True)

        if show_all_years_table:
            st.subheader("年度表格（Y1..Y10，像Excel）")
            st.dataframe(
                df_t[
                    [
                        "year",
                        "owner_spend_emc",
                        "owner_spend_laas",
                        "owner_save_emc",
                        "owner_save_laas",
                        "laas_fee",
                    ]
                ],
                use_container_width=True,
            )

    with tab_provider:
        st.subheader("C) 服务商视角：回报指标（对应 01_Dashboard）")
        c1, c2 = st.columns(2)
        with c1:
            st.metric("NPV（EMC）", _money(dsel.get("emc_npv")))
            st.metric("IRR（EMC）", _pct(dsel.get("emc_irr")))
        with c2:
            st.metric("NPV（LaaS）", _money(dsel.get("laas_npv")))
            st.metric("IRR（LaaS）", _pct(dsel.get("laas_irr")))

        fig = go.Figure()
        fig.add_trace(go.Bar(name="EMC", x=["NPV", "IRR"], y=[dsel.get("emc_npv") or 0.0, (dsel.get("emc_irr") or 0.0) * 1e7], marker_color="#94A3B8"))
        fig.add_trace(go.Bar(name="LaaS", x=["NPV", "IRR"], y=[dsel.get("laas_npv") or 0.0, (dsel.get("laas_irr") or 0.0) * 1e7], marker_color=ACCENT))
        fig.update_layout(
            barmode="group",
            template="plotly_white",
            height=360,
            font_color=NAVY,
            title="服务商KPI对比（IRR为展示缩放）",
        )
        st.plotly_chart(fig, use_container_width=True)

    with tab_cost:
        st.subheader("D) 成本拆解：OPEX包含什么 + CAPEX规模（来自 02_Inputs / 03_Baseline）")
        lamps = dsel.get("lamps") or 0.0
        capex_emc = (dsel.get("capex_emc_per_lamp") or 0.0) * float(lamps)
        capex_laas = (dsel.get("capex_laas_per_lamp") or 0.0) * float(lamps)
        baseline_elec = dsel.get("baseline_electricity_y1")
        if baseline_elec is None or baseline_elec != baseline_elec:
            price = float(dsel.get("electricity_price_per_kwh") or 0.0)
            watts = float(dsel.get("watts_per_lamp") or 0.0)
            hpd = float(dsel.get("hours_per_day") or 0.0)
            dpy = float(dsel.get("days_per_year") or 0.0)
            baseline_elec = float(lamps) * watts / 1000.0 * hpd * dpy * price if (lamps and watts and hpd and dpy and price) else None
        st.markdown(
            f"- **灯具数量**：{_money(lamps)}\n"
            f"- **CAPEX（EMC vs LaaS）**：{_money(capex_emc)} → {_money(capex_laas)}\n"
            f"- **基准电费(Y1)**：{_money(baseline_elec)}\n"
            f"- **节电率（EMC vs LaaS）**：{_pct(dsel.get('emc_saving_rate'))} → {_pct(dsel.get('laas_saving_rate'))}\n"
        )
        opex_breakdown = pd.DataFrame(
            {
                "项目": ["人工/维修（按灯）", "平台费", "备件"],
                "金额": [
                    (dsel.get("opex_om_per_lamp") or 0.0) * float(lamps),
                    dsel.get("opex_platform") or 0.0,
                    dsel.get("opex_spares") or 0.0,
                ],
            }
        )
        fig_o = px.pie(opex_breakdown, names="项目", values="金额", title="非电费运维结构（模板输入拆解）", template="plotly_white")
        fig_o.update_layout(height=360, font_color=NAVY)
        st.plotly_chart(fig_o, use_container_width=True)
        st.caption("说明：电费在模板中通过“基准电费 × (1-节电率)”单独计算，因此此处的运维拆解展示非电费部分。")

    with tab_story:
        st.subheader("E) 为什么更好（证据 + 映射到单元格）")
        cards = cards_for_selected_tier(
            has_upfront=float(dsel.get("upfront") or 0.0) > 0,
            has_tail_discount=float(dsel.get("tail_discount") or 0.0) > 0,
            laas_saving_rate=dsel.get("laas_saving_rate"),
        )
        for c in cards:
            with st.expander(c.title, expanded=(view_mode != "只看关键结论（1分钟版）")):
                st.markdown(f"**为什么**：{c.why}")
                st.markdown(f"**证据/参考**：`{c.evidence_url}`")
                st.markdown("**对应模型单元格**：\n" + "\n".join([f"- `{x}`" for x in c.maps_to_cells]))
                if c.notes:
                    st.markdown(f"**备注**：{c.notes}")

    with tab_trace:
        st.subheader("可追溯性（该方案工作簿的读取范围）")
        td = next((x for x in tiers_dict if x["tier_name"] == sel), None)
        if td is None:
            st.warning("Tier not found in extracted bundle.")
        else:
            st.code(json.dumps(td, ensure_ascii=False, indent=2), language="json")

if __name__ == "__main__":
    main()

