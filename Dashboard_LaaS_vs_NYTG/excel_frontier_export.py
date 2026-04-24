from __future__ import annotations

from dataclasses import asdict
from datetime import date
from pathlib import Path
from typing import Any

import numpy as np
import pandas as pd
import openpyxl
from openpyxl.chart import Reference, ScatterChart, Series
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from business_model_comparison.laas_feasible import (
    ClientValueAssumptions,
    LaaSScenario,
    MAX_AI_OPEX_REDUCTION_PCT,
    MIN_LAST_FOUR_YEAR_FEE_PCT,
    default_fee_grid_from_baseline,
    evaluate_laas_scenario,
    grid_search_feasible_envelope,
)
from business_model_comparison.models import build_baseline_energy_trust
from business_model_comparison.report import (
    baseline_summary_table,
    envelope_table,
    laas_results_to_table,
    provenance_bundle,
    rank_recommended_offers,
)
from business_model_comparison.roadlight_data import load_roadlight_all


ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / "Dashboard_LaaS_vs_NYTG" / "Trust_to_LaaS_Offer_Envelope.xlsx"

# Style palette (matches existing repo navy theme + excel_instruction.md)
NAVY = "1F3864"
TEAL = "0F766E"
LIGHT_GRAY = "F3F4F6"

HEADER_FILL = PatternFill("solid", fgColor=NAVY)
HEADER_FONT = Font(color="FFFFFF", bold=True)
TITLE_FONT = Font(color=NAVY, bold=True, size=18)
SUBTITLE_FONT = Font(color="374151", size=11)
SECTION_FONT = Font(color=NAVY, bold=True, size=12)
BODY_FONT = Font(color="111827", size=10)

WRAP_TOP = Alignment(wrap_text=True, vertical="top")
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)


def _set_col_widths(ws, widths: dict[int, float]) -> None:
    for col, w in widths.items():
        ws.column_dimensions[get_column_letter(col)].width = w


def _header_bar(ws, title: str, subtitle: str) -> None:
    ws.merge_cells("A1:J1")
    ws["A1"] = title
    ws["A1"].font = TITLE_FONT
    ws["A1"].alignment = Alignment(horizontal="left", vertical="center")

    ws.merge_cells("A2:J2")
    ws["A2"] = subtitle
    ws["A2"].font = SUBTITLE_FONT
    ws["A2"].alignment = Alignment(horizontal="left", vertical="center")


def _section(ws, r: int, text: str, c1: int = 1, c2: int = 10) -> int:
    ws.merge_cells(start_row=r, start_column=c1, end_row=r, end_column=c2)
    cell = ws.cell(r, c1, text)
    cell.fill = HEADER_FILL
    cell.font = HEADER_FONT
    cell.alignment = Alignment(horizontal="left", vertical="center")
    return r + 1


def _write_table(ws, start_row: int, start_col: int, df: pd.DataFrame, *, header_fill: PatternFill = HEADER_FILL) -> tuple[int, int]:
    r = start_row
    c = start_col

    # Header row
    for j, col in enumerate(df.columns, start=c):
        cell = ws.cell(r, j, str(col))
        cell.fill = header_fill
        cell.font = HEADER_FONT
        cell.alignment = CENTER
    r += 1

    # Body
    for _, row in df.iterrows():
        for j, col in enumerate(df.columns, start=c):
            val = row[col]
            cell = ws.cell(r, j, val)
            cell.font = BODY_FONT
            cell.alignment = WRAP_TOP
        r += 1

    return r, c + len(df.columns) - 1


def pareto_frontier(df: pd.DataFrame, *, x_col: str, y_col: str) -> pd.DataFrame:
    """Return nondominated points where we maximize y and minimize x.

    Dominance: A dominates B if (y_A >= y_B and x_A <= x_B) and at least one strict.
    """
    d = df[[x_col, y_col] + [c for c in df.columns if c not in (x_col, y_col)]].copy()
    d = d.replace([np.inf, -np.inf], np.nan).dropna(subset=[x_col, y_col])
    d = d.sort_values([x_col, y_col], ascending=[True, False]).reset_index(drop=True)

    frontier_idx: list[int] = []
    best_y = None
    for i, row in d.iterrows():
        x = float(row[x_col])
        y = float(row[y_col])
        if best_y is None or y > best_y + 1e-9:
            frontier_idx.append(int(i))
            best_y = y
        else:
            # dominated by a previous point with <=x and >=y
            continue
    return d.loc[frontier_idx].reset_index(drop=True)


def _top_n_diversified(
    df: pd.DataFrame,
    *,
    n: int = 9,
    min_per_mode: int = 2,
    pct_cap: float = 0.40,
) -> pd.DataFrame:
    """Pick Top-N rows by provider NPV with diversification.

    - Enforce at least `min_per_mode` per opex_mode where possible.
    - Cap ai_opex_reduction_pct for pct modes (uniform_pct / electricity_only_pct) to avoid only 100% solutions.
    - Allow ai_plus_solar separately (cap not applied).
    """
    if df.empty:
        return df

    d = df.copy()
    d["is_pct_mode"] = d["opex_mode"].isin(["uniform_pct", "electricity_only_pct"])
    d = d[(~d["is_pct_mode"]) | (d["ai_opex_reduction_pct"] <= float(pct_cap) + 1e-9)]

    # Rank by provider NPV
    d = d.sort_values("npv_project_rmb", ascending=False)

    picked = []
    for mode in ["uniform_pct", "electricity_only_pct", "ai_plus_solar"]:
        dm = d[d["opex_mode"] == mode].head(int(min_per_mode))
        picked.append(dm)
    base = pd.concat(picked, axis=0).drop_duplicates()

    remaining = d[~d.index.isin(base.index)].head(max(0, int(n) - len(base)))
    out = pd.concat([base, remaining], axis=0).drop_duplicates().head(int(n)).reset_index(drop=True)
    return out


def _summary_block(
    baseline_fee_rmb_y1: float,
    baseline_gp_rmb_y1: float,
    baseline_payback_months: Any,
    offer: dict[str, Any],
) -> pd.DataFrame:
    """Build a screenshot-style comparison block (values in 万元 unless noted)."""
    # Convert RMB → 万元 for display
    base_fee_10k = float(baseline_fee_rmb_y1) / 10_000.0
    rec_fee_10k = float(offer.get("net_client_payment_y1_rmb", offer["annual_service_fee_rmb"])) / 10_000.0
    owner_save_10k = base_fee_10k - rec_fee_10k

    base_gp_10k = float(baseline_gp_rmb_y1) / 10_000.0
    rec_gp_10k = float(offer.get("gp_year1_rmb", 0.0)) / 10_000.0

    base_pb_y = (float(baseline_payback_months) / 12.0) if isinstance(baseline_payback_months, int) else None
    rec_pb_y = (float(offer["payback_months"]) / 12.0) if isinstance(offer["payback_months"], int) else None

    rows = [
        ("业主年度总支出（万元）", base_fee_10k, rec_fee_10k, rec_fee_10k - base_fee_10k, "业主", "下降" if owner_save_10k > 0 else "上升"),
        ("业主年度节约（万元）", 0.0, owner_save_10k, owner_save_10k, "业主", "放大收益" if owner_save_10k > 0 else "减少"),
        ("华普年毛空间（万元）", base_gp_10k, rec_gp_10k, rec_gp_10k - base_gp_10k, "华普", "提升" if rec_gp_10k >= base_gp_10k else "下降"),
        ("华普静态回本周期（年）", base_pb_y if base_pb_y is not None else "NA", rec_pb_y if rec_pb_y is not None else "NA", (rec_pb_y - base_pb_y) if (base_pb_y is not None and rec_pb_y is not None) else "NA", "华普", "缩短" if (base_pb_y is not None and rec_pb_y is not None and rec_pb_y < base_pb_y) else "延长"),
        ("华普NPV（万元）", float(offer.get("baseline_npv_rmb", 0.0)) / 10_000.0, float(offer["npv_project_rmb"]) / 10_000.0, (float(offer["npv_project_rmb"]) - float(offer.get("baseline_npv_rmb", 0.0))) / 10_000.0, "华普", "提升"),
        ("合同定位", "能源托管", f"AI-LaaS（{offer['opex_mode']}）", "重构", "双方", "重构"),
    ]
    df = pd.DataFrame(rows, columns=["结论", "原能源托管", "AI-LaaS推荐方案", "变化", "影响对象", "判断"])
    df["备注"] = ""
    return df


def _select_representative_points(frontier: pd.DataFrame, n: int = 8) -> pd.DataFrame:
    if frontier.empty:
        return frontier
    # Pick points across client_gap quantiles.
    frontier = frontier.sort_values("client_gap_rmb").reset_index(drop=True)
    if len(frontier) <= n:
        return frontier
    qs = np.linspace(0, 1, n)
    idx = sorted({int(round(q * (len(frontier) - 1))) for q in qs})
    return frontier.loc[idx].reset_index(drop=True)


def build_workbook() -> Path:
    # Defaults aligned with Streamlit page but export is deterministic.
    horizon_years = 10
    payback_constraint_months = 36
    discount_rate = 0.12

    fee_low = 0.35
    fee_high = 1.20
    fee_steps = 60

    last_four_year_fee_reduction_grid = [0.0, 200_000.0, 400_000.0, 600_000.0, 800_000.0, 1_000_000.0, 1_200_000.0, 1_500_000.0]
    upfront_grid = [0.0, 200_000.0, 500_000.0, 1_000_000.0, 2_000_000.0]
    opex_modes = ["uniform_pct", "electricity_only_pct", "ai_plus_solar"]
    reduction_grid = [0.0, 0.10, 0.20, 0.30, 0.40, 0.60, 0.80, 0.85]

    client_value = ClientValueAssumptions(
        baseline_outage_hours_per_year=30.0,
        laas_guaranteed_outage_hours_per_year=5.0,
        outage_cost_rmb_per_hour=10_000.0,
        sla_credit_share_to_client=1.0,
        client_discount_rate_annual=0.12,
    )

    parsed = load_roadlight_all(ROOT / "data")
    baseline = build_baseline_energy_trust(parsed, analysis_years=horizon_years, discount_rate_annual=discount_rate)
    fee_grid = default_fee_grid_from_baseline(baseline, pct_low=fee_low, pct_high=fee_high, steps=fee_steps)

    env = grid_search_feasible_envelope(
        baseline,
        term_years=list(range(1, horizon_years + 1)),
        annual_fee_rmb_grid=fee_grid,
        last_four_year_fee_reduction_rmb_grid=last_four_year_fee_reduction_grid,
        upfront_rmb_grid=upfront_grid,
        ai_opex_reduction_grid=reduction_grid,
        discount_rate_annual=discount_rate,
        opex_modes=opex_modes,  # type: ignore[arg-type]
        client_value=client_value,
    )
    env_df = envelope_table(env)

    # Enforce payback constraint as in dashboard filtering
    env_df["payback_ok"] = env_df["payback_months"].apply(lambda x: isinstance(x, int) and int(x) <= int(payback_constraint_months))

    provider_df = env_df[(env_df["provider_feasible"]) & (env_df["payback_ok"])].copy()
    everyone_df = env_df[(env_df["feasible_everyone_better_off"]) & (env_df["payback_ok"])].copy()

    # Add baseline NPV for delta columns in summary
    env_df["baseline_npv_rmb"] = float(baseline.npv_project_rmb)

    # Pareto frontier on provider-feasible universe
    frontier = pareto_frontier(provider_df, x_col="client_gap_rmb", y_col="npv_project_rmb")
    rep = _select_representative_points(frontier, n=8)

    # Top-N diversified everyone-feasible offers
    recommended_df = rank_recommended_offers(everyone_df)
    topN = _top_n_diversified(recommended_df, n=9, min_per_mode=2, pct_cap=MAX_AI_OPEX_REDUCTION_PCT)
    best = None if topN.empty else rank_recommended_offers(topN).head(1).iloc[0].to_dict()
    best_result = None
    if best is not None:
        best_result = evaluate_laas_scenario(
            baseline,
            LaaSScenario(
                term_years=int(best["term_years"]),
                annual_service_fee_rmb=float(best["annual_service_fee_rmb"]),
                last_four_year_fee_reduction_rmb=float(best["last_four_year_fee_reduction_rmb"]),
                upfront_rmb=float(best["upfront_rmb"]),
                ai_opex_reduction_pct=float(best["ai_opex_reduction_pct"]),
                opex_mode=str(best["opex_mode"]),
            ),
            discount_rate_annual=discount_rate,
            client_value=client_value,
        )

    # Build workbook
    wb = openpyxl.Workbook()
    # Remove default sheet
    wb.remove(wb.active)

    # 1) Cover / Dashboard
    ws = wb.create_sheet("Dashboard", 0)
    _set_col_widths(ws, {1: 24, 2: 20, 3: 20, 4: 20, 5: 20, 6: 18, 7: 18, 8: 18, 9: 18, 10: 18})
    _header_bar(
        ws,
        "能源托管 → LaaS | Offer Envelope Dashboard",
        f"Roadlight case | Generated {date.today().isoformat()} | All figures trace to /data + explicit assumptions",
    )

    r = 4
    r = _section(ws, r, "Key KPIs (snapshot)")
    kpis = [
        ("Baseline CAPEX (RMB)", float(baseline.capex_y0_rmb)),
        ("Baseline payback (months)", str(baseline.payback_months)),
        ("Provider-feasible offers (#)", int(provider_df.shape[0])),
        ("Everyone-feasible offers (#)", int(everyone_df.shape[0])),
    ]
    for i, (k, v) in enumerate(kpis):
        rr = r + (i // 2)
        cc = 1 + (i % 2) * 5
        ws.merge_cells(start_row=rr, start_column=cc, end_row=rr, end_column=cc + 1)
        ws.cell(rr, cc, k).font = SECTION_FONT
        ws.cell(rr, cc, k).alignment = Alignment(horizontal="left")
        ws.merge_cells(start_row=rr, start_column=cc + 2, end_row=rr, end_column=cc + 4)
        val_cell = ws.cell(rr, cc + 2, v)
        val_cell.font = Font(color="111827", bold=True, size=12)
        val_cell.alignment = Alignment(horizontal="left")
        if isinstance(v, (int, float)):
            val_cell.number_format = "#,##0"
    r += 3

    r = _section(ws, r, "Best everyone-feasible offer (savings + GP uplift + faster payback)", c1=1, c2=10)
    if best is None:
        ws.cell(r, 1, "No everyone-feasible offers found under constraints.").font = BODY_FONT
        r += 2
    else:
        summary_rows = [
            ("Term (years)", int(best["term_years"])),
            ("Annual service fee (RMB)", float(best["annual_service_fee_rmb"])),
            ("Last 4 years reduction (RMB/year)", float(best["last_four_year_fee_reduction_rmb"])),
            ("Upfront (RMB)", float(best["upfront_rmb"])),
            ("OPEX mode", str(best["opex_mode"])),
            ("Reduction (pct modes)", float(best["ai_opex_reduction_pct"])),
            ("Payback (months)", int(best["payback_months"])),
            ("Payback improvement vs baseline (months)", float(best["payback_improvement_months"])),
            ("Min client savings / year (RMB)", float(best["min_client_savings_rmb_per_year"])),
            ("Min provider GP uplift / year (RMB)", float(best["min_provider_gross_profit_uplift_rmb_per_year"])),
            ("Provider NPV (RMB)", float(best["npv_project_rmb"])),
            ("Client gap (PV RMB)", float(best["client_gap_rmb"])),
        ]
        for i, (k, v) in enumerate(summary_rows):
            ws.cell(r + i, 1, k).font = BODY_FONT
            ws.cell(r + i, 2, v).font = Font(color="111827", bold=True, size=10)
            ws.cell(r + i, 2).number_format = "#,##0" if isinstance(v, (int, float)) else "General"
        r += len(summary_rows) + 1

    # Screenshot-style comparison table (baseline vs best)
    if best is not None:
        best["gp_year1_rmb"] = float(best_result.provider_accounting_gross_profit_rmb_y.get(1)) if best_result is not None else 0.0
        best["net_client_payment_y1_rmb"] = float(best_result.client_payment_rmb_y.get(1)) if best_result is not None else float(best["annual_service_fee_rmb"])
        best["baseline_npv_rmb"] = float(baseline.npv_project_rmb)
        r = _section(ws, r, "Summary table (baseline vs 推荐方案)", c1=1, c2=10)
        summ = _summary_block(
            baseline_fee_rmb_y1=float(baseline.revenue_rmb_y.get(1)),
            baseline_gp_rmb_y1=float(baseline.accounting_gross_profit_rmb_y.get(1)),
            baseline_payback_months=baseline.payback_months,
            offer=best,
        )
        end_r, end_c = _write_table(ws, r, 1, summ)
        r = end_r + 1
        if best_result is not None:
            r = _section(ws, r, "Best offer yearly cashflow analysis (RMB/year)", c1=1, c2=10)
            best_yearly = laas_results_to_table(best_result, baseline)
            end_r, end_c = _write_table(ws, r, 1, best_yearly)
            r = end_r + 1

    # Add scatter chart: provider-feasible offers (client_gap vs NPV), with frontier highlighted via separate table.
    r = _section(ws, r, "Chart: Provider NPV vs Client Gap (provider-feasible universe)", c1=1, c2=10)
    # Put a compact data block for charting.
    # IMPORTANT: take both sides of client_gap around 0 to avoid showing only very negative points.
    chart_df = provider_df[["client_gap_rmb", "npv_project_rmb"]].copy()
    chart_df = chart_df.sort_values("client_gap_rmb")
    neg = chart_df[chart_df["client_gap_rmb"] <= 0].tail(1000)
    pos = chart_df[chart_df["client_gap_rmb"] > 0].head(1000)
    chart_df = pd.concat([neg, pos], axis=0).drop_duplicates()
    chart_df.columns = ["client_gap_rmb", "provider_npv_rmb"]
    end_r, end_c = _write_table(ws, r, 1, chart_df)

    chart = ScatterChart()
    chart.title = "Provider NPV vs Client Gap (smaller is better for client)"
    chart.x_axis.title = "Client gap (PV RMB)  (<=0 means client better off)"
    chart.y_axis.title = "Provider NPV (RMB)"
    chart.legend = None

    xvalues = Reference(ws, min_col=1, min_row=r + 1, max_row=end_r - 1)
    yvalues = Reference(ws, min_col=2, min_row=r + 1, max_row=end_r - 1)
    series = Series(yvalues, xvalues, title="Offers")
    series.marker.symbol = "circle"
    series.marker.size = 4
    chart.series.append(series)
    ws.add_chart(chart, "E14")

    # 2) Executive Summary
    ws2 = wb.create_sheet("Executive_Summary", 1)
    _set_col_widths(ws2, {1: 22, 2: 90, 3: 16})
    _header_bar(ws2, "Executive Summary", "Transitioning 能源托管 → LaaS: feasible offer envelope and trade-offs")
    rr = 4
    rr = _section(ws2, rr, "Objective", c1=1, c2=3)
    ws2.cell(rr, 1, "Objective").font = SECTION_FONT
    ws2.cell(rr, 2, "Find LaaS commercial terms that make the client pay less each year, improve 华普 gross profit each year, and shorten provider payback versus baseline, using a traceable value model and explicit OPEX transformation modes.").alignment = WRAP_TOP
    rr += 2
    rr = _section(ws2, rr, "Key constraints (hard filters)", c1=1, c2=3)
    bullets = [
        "Term ≤ 10 years",
        "Provider simple cash payback ≤ 36 months",
        "Provider payback faster than baseline",
        "Provider accounting gross profit ≥ baseline each year",
        "Client pays less than baseline each year",
        "Client benefit: client_gap ≤ 0, where client_gap = PV(LaaS payments incl upfront) − PV(baseline payments) − PV(ValueFromGuarantees)",
    ]
    for i, btxt in enumerate(bullets):
        ws2.cell(rr + i, 2, f"- {btxt}").alignment = WRAP_TOP
    rr += len(bullets) + 1

    rr = _section(ws2, rr, "Recommendation (commercially aligned ranking)", c1=1, c2=3)
    if best is None:
        ws2.cell(rr, 2, "No everyone-feasible offers under current assumptions. Review client value assumptions, OPEX modes, or allow different commercial structures.").alignment = WRAP_TOP
    else:
        ws2.cell(rr, 2, f"Select an offer that first maximizes minimum annual client savings and minimum annual 华普 gross profit uplift, then prefers faster payback and higher NPV. Current recommendation: term={int(best['term_years'])}y, fee=RMB {best['annual_service_fee_rmb']:,.0f}/y, upfront=RMB {best['upfront_rmb']:,.0f}, mode={best['opex_mode']}.").alignment = WRAP_TOP

    # 3) Assumptions
    ws3 = wb.create_sheet("Assumptions", 2)
    _set_col_widths(ws3, {1: 34, 2: 22, 3: 70})
    _header_bar(ws3, "Assumptions & Inputs", "All assumptions are explicit; source-driven inputs trace to /data")
    rr = 4
    rr = _section(ws3, rr, "Search ranges")
    rows = [
        ("Horizon (years)", horizon_years, ""),
        ("Payback constraint (months)", payback_constraint_months, ""),
        ("Discount rate (annual)", discount_rate, "Used for provider NPV and default client discount rate"),
        ("Service fee low (% baseline)", fee_low, "Fee grid is baseline 年托管费 × pct_low..pct_high"),
        ("Service fee high (% baseline)", fee_high, ""),
        ("Fee steps", fee_steps, ""),
        ("Last 4 years reduction grid (RMB/year)", ", ".join(f"{x:,.0f}" for x in last_four_year_fee_reduction_grid), f"Applied in years 7-10 only; commercial fee floor is {int(MIN_LAST_FOUR_YEAR_FEE_PCT * 100)}% of main annual fee"),
        ("Upfront grid (RMB)", ", ".join(f"{x:,.0f}" for x in upfront_grid), "Client pays at month 0; included in client PV cost"),
        ("OPEX modes", ", ".join(opex_modes), "ai_plus_solar sets electricity OPEX=0 (assumption; no solar CAPEX modeled)"),
        ("Reduction grid (pct)", ", ".join(f"{int(x*100)}%" for x in reduction_grid), f"Used for pct modes; percentage-style AI OPEX reduction capped at {int(MAX_AI_OPEX_REDUCTION_PCT * 100)}%; ai_plus_solar ignores pct grid"),
    ]
    for i, (k, v, note) in enumerate(rows):
        ws3.cell(rr + i, 1, k).font = BODY_FONT
        ws3.cell(rr + i, 2, v).font = Font(color="111827", bold=True, size=10)
        ws3.cell(rr + i, 3, note).alignment = WRAP_TOP
    rr += len(rows) + 2
    rr = _section(ws3, rr, "Client value assumptions (SLA / risk transfer)")
    for i, (k, v) in enumerate(asdict(client_value).items()):
        ws3.cell(rr + i, 1, k).font = BODY_FONT
        ws3.cell(rr + i, 2, v).font = Font(color="111827", bold=True, size=10)
    rr += len(asdict(client_value)) + 2
    rr = _section(ws3, rr, "Traceability notes")
    ws3.cell(rr, 1, "Baseline 托管费 source").font = BODY_FONT
    ws3.cell(rr, 2, "data/income_analysis.csv : 托管收入").alignment = WRAP_TOP
    ws3.cell(rr + 1, 1, "Cash OPEX source").font = BODY_FONT
    ws3.cell(rr + 1, 2, "data/opex.csv : 改造后电费 + 职工薪酬费用 + 维修材料费 + 车辆费用 + 管理费用").alignment = WRAP_TOP
    ws3.cell(rr + 2, 1, "CAPEX source").font = BODY_FONT
    ws3.cell(rr + 2, 2, "data/capex.csv : 总投资").alignment = WRAP_TOP

    # 4) Outputs (Pareto frontier + recommended points)
    ws4 = wb.create_sheet("Outputs", 3)
    _set_col_widths(ws4, {1: 18, 2: 18, 3: 16, 4: 16, 5: 18, 6: 16, 7: 14, 8: 14, 9: 14, 10: 14})
    _header_bar(ws4, "Outputs", "Offer universe, Pareto frontier, and recommended combinations")
    rr = 4
    rr = _section(ws4, rr, "Pareto frontier (provider-feasible points; maximize NPV, minimize client_gap)")
    frontier_show = frontier[
        [
            "term_years",
            "annual_service_fee_rmb",
            "last_four_year_fee_reduction_rmb",
            "upfront_rmb",
            "opex_mode",
            "ai_opex_reduction_pct",
            "payback_months",
            "npv_project_rmb",
            "client_gap_rmb",
            "client_benefit_pass",
            "min_client_savings_rmb_per_year",
            "min_provider_gross_profit_uplift_rmb_per_year",
        ]
    ].copy()
    end_r, end_c = _write_table(ws4, rr, 1, frontier_show.head(200))
    rr = end_r + 2
    rr = _section(ws4, rr, "Representative frontier points (for discussion)")
    rep_show = rep[
        [
            "term_years",
            "annual_service_fee_rmb",
            "last_four_year_fee_reduction_rmb",
            "upfront_rmb",
            "opex_mode",
            "ai_opex_reduction_pct",
            "payback_months",
            "npv_project_rmb",
            "client_gap_rmb",
            "min_client_savings_rmb_per_year",
            "min_provider_gross_profit_uplift_rmb_per_year",
        ]
    ].copy()
    _write_table(ws4, rr, 1, rep_show)

    # 5) Appendix / Backup
    ws5 = wb.create_sheet("Appendix_Data", 4)
    _set_col_widths(ws5, {1: 14, 2: 18, 3: 14, 4: 14, 5: 16, 6: 16, 7: 16, 8: 16, 9: 16, 10: 18, 11: 18, 12: 18})
    _header_bar(ws5, "Appendix / Backup calculations", "Baseline yearly table + offer universe sample + provenance bundle")
    rr = 4
    rr = _section(ws5, rr, "Baseline yearly table (RMB/year)")
    base_df = baseline_summary_table(baseline)
    _write_table(ws5, rr, 1, base_df)
    rr += len(base_df) + 4
    if best_result is not None:
        rr = _section(ws5, rr, "Best offer yearly cashflow analysis (RMB/year)")
        best_df = laas_results_to_table(best_result, baseline)
        _write_table(ws5, rr, 1, best_df)
        rr += len(best_df) + 4
    rr = _section(ws5, rr, "Offer universe sample (first 2,000 rows)")
    sample = env_df.head(2000)
    _write_table(ws5, rr, 1, sample)

    # Traceability JSON as text
    rr = 4
    ws6 = wb.create_sheet("Traceability", 5)
    _set_col_widths(ws6, {1: 24, 2: 110})
    _header_bar(ws6, "Traceability", "Provenance bundle (sources + transforms) for baseline and best offer")
    bundle = provenance_bundle(baseline, best_result)
    rr = 4
    rr = _section(ws6, rr, "Baseline provenance (JSON)", c1=1, c2=2)
    ws6.cell(rr, 1, "baseline").font = SECTION_FONT
    ws6.cell(rr, 2, str(bundle.get("baseline", {}))).alignment = WRAP_TOP

    # Finishing touches
    for sheet in wb.worksheets:
        sheet.sheet_view.showGridLines = False
        sheet.freeze_panes = "A3"

    OUT.parent.mkdir(parents=True, exist_ok=True)
    wb.save(OUT)
    return OUT


def main() -> None:
    out = build_workbook()
    print(f"Wrote Excel: {out}")


if __name__ == "__main__":
    main()

