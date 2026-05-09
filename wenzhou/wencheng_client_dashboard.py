"""
文成隧道及路灯改造 — 客户侧经济性测算（独立 Streamlit 应用逻辑）。

数据：本目录下 wencheng_retrofit_and_om_baseline.csv

运行（自仓库根目录，仅此单页）：
  python3 -m streamlit run wenzhou/run_dashboard.py
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
def load_wencheng_csv(path_str: str, file_mtime_ns: int) -> pd.DataFrame:
    """mtime 参与缓存键：CSV 重新生成后自动失效，避免仍显示旧表尾/改造后电量。"""
    del file_mtime_ns  # 仅用于缓存失效
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
    """`agg_savings_fee_total`：节费（元/年）= 表尾改造后电量×λ = 改造前电费−分项节电量合计×λ；非「分项改造后电费」。"""
    r = row_by_id(df, "agg_savings_fee_total")
    v = _safe_float(r.get("annual_fee_cny"))
    if np.isnan(v):
        v = _safe_float(r.get("ratio_value"))
    return v


def sum_savings_line_annual_fees(df: pd.DataFrame) -> float:
    """分项 savings_line 的年节电费（元）合计；等于节电量合计×各行 λ（通常各行 λ 相同）。"""
    sl = df.loc[df["record_type"] == "savings_line"]
    if sl.empty:
        return float("nan")
    return float(sl["annual_fee_cny"].astype(float).sum())


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


def get_kwh_before_footer(df: pd.DataFrame) -> float:
    return _safe_float(row_by_id(df, "agg_before_kwh_footer").get("ratio_value"))


def get_kwh_after_footer(df: pd.DataFrame) -> float:
    return _safe_float(row_by_id(df, "agg_after_kwh_footer").get("ratio_value"))


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
    *,
    money_round: int = 6,
) -> pd.DataFrame:
    """money_round：内部累计保留小数位，便于与「节费」等三位小数展示对齐。"""
    rows: list[dict[str, float]] = []
    cum = 0.0
    for year in range(0, horizon_years + 1):
        if year == 0:
            net = round(-float(investment), money_round)
        else:
            net = round(float(annual_benefit), money_round)
        cum = round(cum + net, money_round)
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
    csv_mtime_ns = Path(csv_path).stat().st_mtime_ns

    st.markdown(
        f"<h2 style='color:{NAVY};margin-bottom:0.2rem;'>文成隧道路灯改造 · 客户侧经济性</h2>"
        f"<p style='color:{MUTED};margin-top:0;'>测算改造投资、回收期与现金流（简化：年末入账、未折现）。"
        f" 从第 1 年起，客户现金流入 = <strong>年度节费</strong>（节省下来的电费）× <strong>客户分成</strong>；"
        f"其余节费归<strong>业主</strong>。左侧滑块调节该比例。</p>",
        unsafe_allow_html=True,
    )

    try:
        df = load_wencheng_csv(str(csv_path), csv_mtime_ns)
    except FileNotFoundError:
        st.error(f"未找到数据文件：{csv_path}")
        st.stop()

    if has_reconcile_notes(df):
        st.warning(
            "数据中包含 **行汇总与表尾合计不一致** 的核对标记（reconcile_required）。"
            "本页电量均按表尾合计；分项差异请见 CSV。"
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
        # 表尾：改造前/后电量；「改造后电费」业务口径 = 分项节电量合计×λ（≈483,636）；节费 = 改造前−该项 = 表尾改造后电量×λ（≈412,618）
        kwh_before = get_kwh_before_footer(df)
        kwh_after = get_kwh_after_footer(df)
        original_annual_fee = round(float(kwh_before) * float(tariff), 6)
        bill_post_lines_cny = round(float(kwh_saved_default) * float(tariff), 6)
        annual_fee_after_retrofit_footer = round(float(kwh_after) * float(tariff), 6)

        capex_share_pct = st.slider("客户承担改造投资比例（%）", 0, 100, 100, 1)
        benefit_share_pct = st.slider(
            "节费分成：客户比例（%）",
            0,
            100,
            100,
            1,
            help="年度节费 = 改造前电费 −（分项节电量合计×λ）；等于表尾改造后电量×λ；与下方对照表一致；驱动现金流与 IRR。",
        )
        st.caption(f"业主分得：**{100 - benefit_share_pct}%**（节费剩余部分）")
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
        st.caption("启动命令（仅此一页；入口在 wenzhou/ 下，不会加载仓库根 pages/）")
        st.code("python3 -m streamlit run wenzhou/run_dashboard.py", language="bash")

    # 节费（现金流池）= 改造前电费 −（分项：节电量合计×λ）= 表尾改造后电量×λ。「分项节电量×λ」= 收资表「改造后电费」合计（≈483,636）；节费 ≈412,618。
    if (
        np.isfinite(original_annual_fee)
        and np.isfinite(bill_post_lines_cny)
        and np.isfinite(annual_fee_after_retrofit_footer)
    ):
        annual_shouru = round(
            float(original_annual_fee) - float(bill_post_lines_cny), 6
        )
    elif np.isfinite(annual_fee_after_retrofit_footer):
        annual_shouru = annual_fee_after_retrofit_footer
    else:
        annual_shouru = float("nan")

    line_fees_csv = sum_savings_line_annual_fees(df)
    if (
        np.isfinite(line_fees_csv)
        and np.isfinite(tariff_default)
        and abs(float(tariff) - float(tariff_default)) < 1e-9
        and abs(float(line_fees_csv) - float(bill_post_lines_cny)) > 1.0
    ):
        st.warning(
            f"分项 `savings_line` 年节电费合计 **{line_fees_csv:,.3f}** 与 节电量合计×λ **{bill_post_lines_cny:,.3f}** 相差超过 1 元，请核对 CSV。"
        )

    capex_share = capex_share_pct / 100.0
    benefit_share = benefit_share_pct / 100.0

    client_investment = round(float(total_capex_default) * capex_share + float(extra_cost), 6)
    # 客户现金流 = 年度节费 × 客户分成（节费需分一部分给业主）
    owner_annual_savings_share = round(float(annual_shouru) * (1.0 - benefit_share), 6)
    client_annual_cashflow_in = round(float(annual_shouru) * benefit_share, 6)

    if (not np.isfinite(annual_shouru)) or annual_shouru <= 0:
        st.warning(
            "**节费（改造前电费 − 分项节电量×λ）** 当前 ≤ 0 或无法计算，客户/业主分成与现金流无意义。"
            "请检查 CSV 表尾电量、`agg_savings_kwh_total` 与电价 λ。"
        )
    elif (
        np.isfinite(annual_fee_after_retrofit_footer)
        and np.isfinite(annual_shouru)
        and abs(float(annual_fee_after_retrofit_footer) - float(annual_shouru)) > 1.0
    ):
        st.warning(
            f"表尾改造后电量×λ（**{annual_fee_after_retrofit_footer:,.3f}**）与「改造前−分项节电量×λ」节费 **{annual_shouru:,.3f}** "
            "相差超过 1 元；请核对表尾与 agg_savings_kwh_total。"
        )

    # --- 电费与年收入（说明 + 对照表）---
    st.subheader("电费对照 · 一句话看懂")
    st.markdown(
        """
**对照表**  
- **改造前**：表尾改造前年电量 × λ → 改造前年电费（例 **896,254.56** 元/年 @ λ=0.72）。  
- **改造后电费（分项口径）**：收资表「改造后电费」列 = **节电量合计 × λ**（例 **483,636.096** 元/年）——**不是**表尾「改造后年用电量×λ」。  
- **年节费（少付、现金流基数）** = 改造前年电费 **−** 上一项 = **表尾改造后年用电量 × λ**（例 **412,618.464** 元/年）。  

**现金流** = **年节费** × **客户分成**。调节 λ 时，第 2、3 行与节费联动；CSV 分项「年节电费」列仅在 λ 与文件一致时可与第 2 行逐字核对。  
"""
    )
    fee_compare_tbl = pd.DataFrame(
        {
            "项目": [
                "改造前（表尾旧工况）",
                "改造后电费（分项：节电量合计×λ）",
                "年节费（少付）= 表尾改造后电量×λ",
            ],
            "年电量（kWh）": [
                f"{kwh_before:,.2f}" if np.isfinite(kwh_before) else "—",
                f"{kwh_saved_default:,.2f}" if np.isfinite(kwh_saved_default) else "—",
                f"{kwh_after:,.2f}" if np.isfinite(kwh_after) else "—",
            ],
            "年电费（元/年）": [
                f"{original_annual_fee:,.3f}" if np.isfinite(original_annual_fee) else "—",
                f"{bill_post_lines_cny:,.3f}" if np.isfinite(bill_post_lines_cny) else "—",
                f"{annual_shouru:,.3f}" if np.isfinite(annual_shouru) else "—",
            ],
        }
    )
    st.dataframe(fee_compare_tbl, use_container_width=True, hide_index=True)
    with st.expander("展开：分项「改造后电费」与表尾节费（与收资表一致）", expanded=False):
        st.markdown(
            "**改造后电费（分项）** = 节电量合计 × λ，对应收资表「改造后电费」列合计。"
            "**年节费** = 改造前电费 − 该项 = **表尾改造后年用电量 × λ**（现金流分成基数）。"
        )
        st.dataframe(
            pd.DataFrame(
                {
                    "说明": [
                        "改造后电费（分项：节电量合计×λ）",
                        "年节费（表尾改造后电量×λ）",
                    ],
                    "年电量（kWh）": [
                        f"{kwh_saved_default:,.2f}" if np.isfinite(kwh_saved_default) else "—",
                        f"{kwh_after:,.2f}" if np.isfinite(kwh_after) else "—",
                    ],
                    "年电费（元/年）": [
                        f"{bill_post_lines_cny:,.3f}" if np.isfinite(bill_post_lines_cny) else "—",
                        f"{annual_fee_after_retrofit_footer:,.3f}"
                        if np.isfinite(annual_fee_after_retrofit_footer)
                        else "—",
                    ],
                }
            ),
            use_container_width=True,
            hide_index=True,
        )
    st.caption(
        f"λ = **{tariff:.2f}** 元/kWh；第 1、3 行电量为 **表尾**；第 2 行为 **分项节电量合计**。"
        " 现金流基数为 **节费**（第 3 行）。"
    )
    sp1, sp2, sp3 = st.columns(3)
    sp1.metric(
        "节费总额（可分配，元/年）",
        f"{annual_shouru:,.3f}" if np.isfinite(annual_shouru) else "—",
        help="改造前年电费 −（分项节电量合计×λ）；等于表尾改造后电量×λ；客户与业主从该池分成。",
    )
    sp2.metric(
        f"其中 — 客户（{benefit_share_pct}%）",
        f"{client_annual_cashflow_in:,.3f}" if np.isfinite(client_annual_cashflow_in) else "—",
        help="计入下方现金流、回收期与 IRR。",
    )
    sp3.metric(
        f"其中 — 业主（{100 - benefit_share_pct}%）",
        f"{owner_annual_savings_share:,.3f}" if np.isfinite(owner_annual_savings_share) else "—",
        help="节费中归业主的部分（并列展示，不计入客户现金流表）。",
    )

    simple_payback = fractional_payback_years(client_investment, client_annual_cashflow_in)
    cumulative_net = round(
        -float(client_investment) + float(client_annual_cashflow_in) * float(horizon), 6
    )
    simple_roi = (cumulative_net / client_investment) if client_investment > 0 else float("nan")

    flows = [-client_investment] + [client_annual_cashflow_in] * int(horizon)
    irr_val = irr_annual(flows)

    cash_df = compute_cashflow_table(client_investment, client_annual_cashflow_in, int(horizon))
    payback_year_int = first_integer_year_payback(cash_df)

    # --- KPI 指标条 ---
    c1, c2, c3, c4, c5, c6 = st.columns(6)
    c1.metric(
        "客户承担投资额（元）",
        f"{client_investment:,.0f}",
        help="改造总投资 × 客户承担比例 + 额外一次性支出",
    )
    c2.metric(
        "年均现金流收入—归属客户（元/年）",
        f"{client_annual_cashflow_in:,.3f}",
        help="年度节费 × 客户分成比例；与左侧滑块一致；计入现金流表。",
    )
    c3.metric(
        "简单回收期（年）",
        "—" if simple_payback is None else f"{simple_payback:.2f}",
        help="客户总投资 ÷ 年均现金流收入（未折现）",
    )
    c4.metric(
        f"分析期（{int(horizon)}年）累计净收益（元）",
        f"{cumulative_net:,.3f}",
        help="− 总投资 + 年均现金流收入 × 年数（简化）",
    )
    c5.metric(
        f"分析期简单收益率",
        "—" if not np.isfinite(simple_roi) else f"{simple_roi * 100:.1f}%",
        help="累计净收益 ÷ 客户总投资",
    )
    c6.metric(
        "年化 IRR",
        "—" if irr_val is None else f"{irr_val * 100:.2f}%",
        help="年度现金流（每年一期）下的 IRR，即年化贴现率；numpy-financial.irr；无解时显示 —",
    )

    if irr_val is not None and np.isfinite(irr_val):
        st.markdown(
            f"<div style='text-align:center;margin:0.6rem 0 0.25rem;padding:14px 20px;"
            f"background:#F8FAFC;border-radius:10px;border:1px solid #E2E8F0;'>"
            f"<div style='font-size:0.82rem;color:{MUTED};letter-spacing:0.02em;'>年化内部收益率 · IRR</div>"
            f"<div style='font-size:2.35rem;font-weight:700;color:{ACCENT};line-height:1.15;"
            f"font-variant-numeric:tabular-nums;'>{irr_val * 100:.2f}%</div>"
            f"<div style='font-size:0.78rem;color:{MUTED};margin-top:6px;'>"
            f"基于上方年度现金流序列（第 0 年投资，随后 {int(horizon)} 年等额流入）</div></div>",
            unsafe_allow_html=True,
        )

    if np.isfinite(kwh_before) and np.isfinite(original_annual_fee) and np.isfinite(bill_post_lines_cny):
        st.caption(
            f"客户现金流：**{annual_shouru:,.3f}** × **{benefit_share_pct}%** = **{client_annual_cashflow_in:,.3f}** 元/年；"
            f" 业主：**{owner_annual_savings_share:,.3f}** 元/年。"
            f" 分项「改造后电费」合计：**{bill_post_lines_cny:,.3f}** 元/年；"
            f" 年节费（表尾）：**{annual_fee_after_retrofit_footer:,.3f}** 元/年。"
        )
    else:
        st.caption("无法完整展示电费拆解：请确认 CSV 中含改造前/后电量聚合行。")

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
                "当年净现金流（元）": st.column_config.NumberColumn(format="%.3f"),
                "累计净现金流（元）": st.column_config.NumberColumn(format="%.3f"),
            },
        )

    # --- 公式说明 ---
    with st.expander("指标计算公式（中文）", expanded=False):
        st.markdown(
            f"""
- **客户承担投资额** = 改造总投资（`agg_capex_total` → **{total_capex_default:,.0f}** 元）× **{capex_share_pct}%** + 额外一次性支出 **{extra_cost:,.0f}** 元  
  → **{client_investment:,.2f}** 元  

- **对照表「改造前」行** = 表尾改造前年电量 × λ → **{original_annual_fee:,.3f}** 元/年  

- **对照表「改造后电费（分项）」行** = 节电量合计 × λ → **{bill_post_lines_cny:,.3f}** 元/年（收资表「改造后电费」列合计）  

- **年节费（现金流基数）** = **{original_annual_fee:,.3f}** − **{bill_post_lines_cny:,.3f}** = **{annual_shouru:,.3f}** 元/年 = 表尾改造后 **{kwh_after:,.2f}** kWh × λ  

- **`agg_savings_fee_total`（CSV）** = **{fee_save_default:,.3f}** 元/年 = `footer_after_kWh` × λ（**节费**，非分项「改造后电费」）  

- **年均现金流收入—归属客户** = 节费 **{annual_shouru:,.3f}** × **{benefit_share_pct}%** → **{client_annual_cashflow_in:,.3f}** 元/年；**业主节费份额** = **{owner_annual_savings_share:,.3f}** 元/年  

- **简单回收期** = 客户承担投资额 ÷ 年均现金流收入（归属客户）

- **分析期累计净收益** = −客户承担投资额 + 年均现金流收入（归属客户） × **{int(horizon)}**

- **分析期简单收益率** = 累计净收益 ÷ 客户承担投资额  

- **年化 IRR**：同上年度现金流序列；**np.irr** 给出的一期利率即「每年一期」下的 **年化收益率**（等价年化）。
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

        kwh_before_footer = _safe_float(row_by_id(df, "agg_before_kwh_footer")["ratio_value"])
        kwh_saved = kwh_saved_default
        denom = kwh_before_footer
        label = "改造前表尾电量（kWh）"
        if denom and denom > 0:
            ratio_save = kwh_saved / denom
            st.metric("节电量 / 改造前表尾", f"{ratio_save * 100:.2f}%", help=f"{label} 来自 CSV。")

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
