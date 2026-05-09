"""
文成隧道及路灯改造 — 客户侧经济性测算（独立 Streamlit 应用逻辑）。

数据：本目录下 wencheng_retrofit_and_om_baseline.csv

运行（自仓库根目录）：
  streamlit run streamlit_wenzhou_client.py
"""

from __future__ import annotations

import sys
from pathlib import Path

import numpy as np
import pandas as pd
import plotly.graph_objects as go
import streamlit as st

# 本文件位于 wenzhou/，CSV 与之一级
_WENZHOU_DIR = Path(__file__).resolve().parent
_CSV_PATH = _WENZHOU_DIR / "wencheng_retrofit_and_om_baseline.csv"
# 相对仓库根的路径（用于界面展示）
_CSV_DISPLAY = Path("wenzhou") / "wencheng_retrofit_and_om_baseline.csv"

_REPO_ROOT = _WENZHOU_DIR.parent
if str(_REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(_REPO_ROOT))

NAVY = "#1F3864"
MUTED = "#6B7280"
ACCENT = "#2563EB"
WARN_ORANGE = "#EA580C"


def _safe_float(v: object, default: float = float("nan")) -> float:
    if v is None or (isinstance(v, float) and np.isnan(v)):
        return default
    if isinstance(v, str) and v.strip() == "":
        return default
    try:
        return float(v)
    except (TypeError, ValueError):
        return default


@st.cache_data(show_spinner=False)
def load_wencheng_csv(path_str: str) -> pd.DataFrame:
    p = Path(path_str)
    if not p.is_file():
        raise FileNotFoundError(str(p))
    return pd.read_csv(p)


def row_by_id(df: pd.DataFrame, row_id: str) -> pd.Series:
    sub = df.loc[df["row_id"] == row_id]
    if sub.empty:
        raise KeyError(f"缺少 row_id={row_id}")
    return sub.iloc[0]


def get_total_capex(df: pd.DataFrame) -> float:
    r = row_by_id(df, "agg_capex_total")
    v = _safe_float(r.get("line_capex_cny"))
    if np.isnan(v):
        v = _safe_float(r.get("ratio_value"))
    return v


def get_annual_fee_saving(df: pd.DataFrame) -> float:
    r = row_by_id(df, "agg_savings_fee_total")
    v = _safe_float(r.get("annual_fee_cny"))
    if np.isnan(v):
        v = _safe_float(r.get("ratio_value"))
    return v


def get_kwh_saved(df: pd.DataFrame) -> float:
    r = row_by_id(df, "agg_savings_kwh_total")
    v = _safe_float(r.get("annual_kwh_saved_line"))
    if np.isnan(v):
        v = _safe_float(r.get("ratio_value"))
    return v


def get_default_tariff(df: pd.DataFrame) -> float:
    sl = df.loc[df["record_type"] == "savings_line"]
    if sl.empty:
        return 0.72
    return float(sl.iloc[0]["tariff_cny_per_kwh"])


def has_reconcile_notes(df: pd.DataFrame) -> bool:
    if "notes" not in df.columns:
        return False
    return bool(df["notes"].fillna("").astype(str).str.contains("reconcile_required").any())


def irr_annual(cashflows: list[float]) -> float | None:
    """年度现金流 IRR；依赖 numpy-financial。"""
    try:
        import numpy_financial as npf  # type: ignore[import-untyped]
    except ImportError:
        return None
    arr = np.array(cashflows, dtype=float)
    try:
        r = npf.irr(arr)
    except Exception:
        return None
    if r is None or (isinstance(r, float) and (np.isnan(r) or np.isinf(r))):
        return None
    return float(r)


def compute_cashflow_table(
    investment: float,
    annual_benefit: float,
    horizon_years: int,
) -> pd.DataFrame:
    rows: list[dict[str, float]] = []
    cum = 0.0
    for year in range(0, horizon_years + 1):
        if year == 0:
            net = -investment
        else:
            net = annual_benefit
        cum += net
        rows.append({"年份": year, "当年净现金流（元）": net, "累计净现金流（元）": cum})
    return pd.DataFrame(rows)


def fractional_payback_years(investment: float, annual_benefit: float) -> float | None:
    if annual_benefit <= 0 or investment < 0:
        return None
    return investment / annual_benefit


def first_integer_year_payback(cash_df: pd.DataFrame) -> int | None:
    """首个累计 >= 0 的年份（整数年）。"""
    for _, r in cash_df.iterrows():
        if r["累计净现金流（元）"] >= 0:
            return int(r["年份"])
    return None


def zh_rename_detail(df: pd.DataFrame) -> pd.DataFrame:
    mapping = {
        "record_type": "记录类型",
        "row_id": "行标识（技术字段）",
        "luminaire_type_zh": "灯具类型",
        "luminaire_type_en": "类型（英文）",
        "power_w": "功率（W）",
        "qty_units": "数量（盏）",
        "hours_per_year": "年点亮小时（h）",
        "annual_kwh_theoretical": "年理论电量（kWh）",
        "delta_power_w": "功率降幅（W）",
        "annual_kwh_saved_line": "年节电量（kWh）",
        "tariff_cny_per_kwh": "综合电价（元/kWh）",
        "annual_fee_cny": "理论年节电费（元）",
        "unit_capex_cny": "单位成本（元/盏）",
        "line_capex_cny": "分项合计（元）",
        "ratio_name": "指标名",
        "ratio_value": "指标值",
        "notes": "备注",
    }
    out = df.copy()
    out = out.rename(columns={k: v for k, v in mapping.items() if k in out.columns})
    return out


def main() -> None:
    st.set_page_config(
        page_title="文成改造 · 客户经济性",
        layout="wide",
        initial_sidebar_state="expanded",
    )

    csv_path = _CSV_PATH

    st.markdown(
        f"<h2 style='color:{NAVY};margin-bottom:0.2rem;'>文成隧道路灯改造 · 客户侧经济性</h2>"
        f"<p style='color:{MUTED};margin-top:0;'>基于调研表汇总 CSV，展示投资额、年收益、回收期、累计现金流与回报率。"
        f" 现金流为<strong>简化模型</strong>（年末一次性、未折现）。</p>",
        unsafe_allow_html=True,
    )

    try:
        df = load_wencheng_csv(str(csv_path))
    except FileNotFoundError:
        st.error(f"未找到数据文件：{csv_path}")
        st.stop()

    if has_reconcile_notes(df):
        st.warning(
            "数据中包含 **行汇总与表尾合计不一致** 的核对标记（reconcile_required）。"
            "电量比例测算时请在下栏选择「表尾合计」或「分项加总」。"
        )

    total_capex_default = get_total_capex(df)
    fee_save_default = get_annual_fee_saving(df)
    kwh_saved_default = get_kwh_saved(df)
    tariff_default = get_default_tariff(df)

    with st.sidebar:
        st.subheader("情景假设")
        tariff = st.number_input(
            "综合电价 λ（元/kWh）",
            min_value=0.0,
            max_value=5.0,
            value=float(tariff_default),
            step=0.01,
            help="默认取自 savings_line；可与「年节电费」交叉校验：节电量 × 电价。",
        )
        use_recomputed_fee = st.checkbox(
            "用电价 × 节电量重算年节电费",
            value=False,
            help="勾选后：年节电费 = 节电量（agg_savings_kwh_total）× 上方电价，否则用 CSV 中的理论年节电费合计。",
        )
        annual_fee_raw = kwh_saved_default * tariff if use_recomputed_fee else fee_save_default

        capex_share_pct = st.slider("客户承担改造投资比例（%）", 0, 100, 100, 1)
        benefit_share_pct = st.slider("客户享有的年节电费比例（%）", 0, 100, 100, 1)
        horizon = st.number_input("分析期（年）", min_value=1, max_value=40, value=10, step=1)
        extra_cost = st.number_input(
            "额外一次性支出（元）",
            min_value=0.0,
            value=0.0,
            step=1000.0,
            help="CSV 未包含的安装费等，计入客户总投资。",
        )

        st.divider()
        st.caption("数据来源文件（相对仓库根）")
        st.code(str(_CSV_DISPLAY), language=None)
        st.caption("启动命令（独立应用，不含其它 multipage）")
        st.code("streamlit run streamlit_wenzhou_client.py", language="bash")

    capex_share = capex_share_pct / 100.0
    benefit_share = benefit_share_pct / 100.0

    client_investment = total_capex_default * capex_share + extra_cost
    client_annual_benefit = annual_fee_raw * benefit_share

    simple_payback = fractional_payback_years(client_investment, client_annual_benefit)
    cumulative_net = -client_investment + client_annual_benefit * float(horizon)
    simple_roi = (cumulative_net / client_investment) if client_investment > 0 else float("nan")

    flows = [-client_investment] + [client_annual_benefit] * int(horizon)
    irr_val = irr_annual(flows)

    cash_df = compute_cashflow_table(client_investment, client_annual_benefit, int(horizon))
    payback_year_int = first_integer_year_payback(cash_df)

    # --- KPI 指标条 ---
    c1, c2, c3, c4, c5, c6 = st.columns(6)
    c1.metric(
        "客户承担投资额（元）",
        f"{client_investment:,.0f}",
        help="改造总投资 × 客户承担比例 + 额外一次性支出",
    )
    c2.metric(
        "年均收益—归属客户（元/年）",
        f"{client_annual_benefit:,.0f}",
        help="年节电费 × 客户享有的收益比例",
    )
    c3.metric(
        "简单回收期（年）",
        "—" if simple_payback is None else f"{simple_payback:.2f}",
        help="客户总投资 ÷ 年均归属收益（未折现）",
    )
    c4.metric(
        f"分析期（{int(horizon)}年）累计净收益（元）",
        f"{cumulative_net:,.0f}",
        help="− 总投资 + 年均收益 × 年数（简化）",
    )
    c5.metric(
        f"分析期简单收益率",
        "—" if not np.isfinite(simple_roi) else f"{simple_roi * 100:.1f}%",
        help="累计净收益 ÷ 客户总投资",
    )
    c6.metric(
        "内部收益率（IRR，年化）",
        "—" if irr_val is None else f"{irr_val * 100:.2f}%",
        help="numpy-financial 基于年度现金流；若无解显示 —",
    )

    # --- 累计现金流图 ---
    st.subheader("累计现金流与回收期")
    fig = go.Figure()
    fig.add_trace(
        go.Scatter(
            x=cash_df["年份"],
            y=cash_df["累计净现金流（元）"],
            mode="lines+markers",
            name="累计净现金流",
            line=dict(color=ACCENT, width=3),
            marker=dict(size=8),
        )
    )
    fig.add_hline(y=0, line_dash="dash", line_color=MUTED, annotation_text="盈亏平衡（0）")
    if simple_payback is not None and np.isfinite(simple_payback) and simple_payback <= horizon:
        fig.add_vline(
            x=simple_payback,
            line_width=2,
            line_dash="dash",
            line_color=WARN_ORANGE,
            annotation_text=f"简单回收期 ≈ {simple_payback:.2f} 年",
            annotation_position="top left",
        )
    elif payback_year_int is not None:
        fig.add_vline(
            x=float(payback_year_int),
            line_width=2,
            line_dash="dot",
            line_color=WARN_ORANGE,
            annotation_text=f"第 {payback_year_int} 年末累计非负",
            annotation_position="top left",
        )

    fig.update_layout(
        template="plotly_white",
        height=420,
        font_color=NAVY,
        xaxis_title="年份（0 年为投资支出）",
        yaxis_title="累计净现金流（元）",
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="left", x=0),
        margin=dict(l=40, r=40, t=60, b=40),
    )
    st.plotly_chart(fig, use_container_width=True)

    with st.expander("查看现金流明细表（与上图一致）", expanded=False):
        st.dataframe(
            cash_df,
            use_container_width=True,
            column_config={
                "年份": st.column_config.NumberColumn("年份", format="%d"),
                "当年净现金流（元）": st.column_config.NumberColumn(format="%.2f"),
                "累计净现金流（元）": st.column_config.NumberColumn(format="%.2f"),
            },
        )

    # --- 公式说明 ---
    with st.expander("指标计算公式（中文）", expanded=False):
        st.markdown(
            f"""
- **客户承担投资额** = 改造总投资（`agg_capex_total` → **{total_capex_default:,.0f}** 元）× **{capex_share_pct}%** + 额外一次性支出 **{extra_cost:,.0f}** 元  
  → **{client_investment:,.2f}** 元  

- **年均归属收益** = 年节电费（{'电价×节电量' if use_recomputed_fee else 'CSV 合计'} → **{annual_fee_raw:,.2f}** 元/年）× **{benefit_share_pct}%**  
  → **{client_annual_benefit:,.2f}** 元/年  

- **简单回收期** = 客户承担投资额 ÷ 年均归属收益  

- **分析期累计净收益** = −客户承担投资额 + 年均归属收益 × **{int(horizon)}**  

- **分析期简单收益率** = 累计净收益 ÷ 客户承担投资额  

- **IRR**：对现金流 [**−投资**, 随后 **{int(horizon)}** 年每年 **+年均归属收益**] 求内部收益率（年化）。
"""
        )

    # --- 成本明细 ---
    st.subheader("成本与明细溯源")
    tab_a, tab_b, tab_c = st.tabs(["改造成本分项", "汇总与比例", "运维假设"])

    with tab_a:
        sl = df[df["record_type"] == "savings_line"].copy()
        if sl.empty:
            st.info("无 savings_line 记录。")
        else:
            disp = sl[
                [
                    "row_id",
                    "qty_units",
                    "hours_per_year",
                    "delta_power_w",
                    "annual_kwh_saved_line",
                    "tariff_cny_per_kwh",
                    "annual_fee_cny",
                    "unit_capex_cny",
                    "line_capex_cny",
                ]
            ].copy()
            disp = zh_rename_detail(disp)
            st.dataframe(disp, use_container_width=True)
            st.caption(
                f"分项 **line_capex_cny** 合计应核对 **{sl['line_capex_cny'].sum():,.0f}** 元；"
                f" 表尾改造总投资 **{total_capex_default:,.0f}** 元（`agg_capex_total`）。"
            )

    with tab_b:
        agg_show = df[df["record_type"] == "aggregate"][
            ["row_id", "ratio_name", "ratio_value", "annual_fee_cny", "annual_kwh_saved_line", "line_capex_cny", "notes"]
        ].copy()
        agg_show = zh_rename_detail(agg_show)
        st.dataframe(agg_show, use_container_width=True)

        baseline = st.radio(
            "节电比例测算基准（表尾合计 vs 分项加总）",
            ["表尾合计（推荐展示）", "分项行加总"],
            horizontal=True,
            help="表尾与分项不一致时，比例不同；详见 CSV 备注 reconcile_required。",
        )
        kwh_before_footer = _safe_float(row_by_id(df, "agg_before_kwh_footer")["ratio_value"])
        kwh_before_lines = _safe_float(row_by_id(df, "agg_before_kwh_sum_lines")["ratio_value"])
        kwh_saved = kwh_saved_default
        if baseline.startswith("表尾"):
            denom = kwh_before_footer
            label = "改造前表尾电量（kWh）"
        else:
            denom = kwh_before_lines
            label = "改造前分项加总电量（kWh）"
        if denom and denom > 0:
            ratio_save = kwh_saved / denom
            st.metric("节电量 / 所选基准", f"{ratio_save * 100:.2f}%", help=f"{label} 来自 CSV。")

    with tab_c:
        om = df[df["record_type"] == "om_assumption"].copy()
        if om.empty:
            st.info("无运维假设记录。")
        else:
            st.dataframe(zh_rename_detail(om), use_container_width=True)

    st.divider()
    st.caption(
        "内部字段 row_id 对应 wenzhou/wencheng_retrofit_and_om_baseline.csv；"
        "英文列名为原始字段名，界面已译为中文。"
    )


if __name__ == "__main__":
    main()
