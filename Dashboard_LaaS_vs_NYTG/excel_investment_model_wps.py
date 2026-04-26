from __future__ import annotations

from datetime import date
from pathlib import Path

import xlsxwriter

from business_model_comparison.models import build_baseline_energy_trust
from business_model_comparison.roadlight_data import load_roadlight_all
import numpy_financial as npf


ROOT = Path(__file__).resolve().parents[1]
OUT_CN = ROOT / "Dashboard_LaaS_vs_NYTG" / "Trust_to_LaaS_Investment_Model_CN_WPS.xlsx"
OUT_EN = ROOT / "Dashboard_LaaS_vs_NYTG" / "Trust_to_LaaS_Investment_Model_EN_WPS.xlsx"
DEBUG_LOG_PATH = ROOT / ".cursor" / "debug-9617d5.log"


# #region agent log
def _dlog(location: str, message: str, data: dict, *, run_id: str = "pre-fix", hypothesis_id: str = "H_wps") -> None:
    import json
    import time

    payload = {
        "sessionId": "9617d5",
        "runId": run_id,
        "hypothesisId": hypothesis_id,
        "location": location,
        "message": message,
        "data": data,
        "timestamp": int(time.time() * 1000),
    }
    try:
        DEBUG_LOG_PATH.parent.mkdir(parents=True, exist_ok=True)
        with DEBUG_LOG_PATH.open("a", encoding="utf-8") as f:
            f.write(json.dumps(payload, ensure_ascii=False) + "\n")
    except Exception:
        return


# #endregion agent log


def _formats(wb: xlsxwriter.Workbook) -> dict[str, xlsxwriter.format.Format]:
    navy = "#1F3864"
    return {
        "title": wb.add_format({"bold": True, "font_size": 16, "font_color": navy}),
        "subtitle": wb.add_format({"font_size": 10, "font_color": "#374151"}),
        "header": wb.add_format({"bold": True, "font_color": "#FFFFFF", "bg_color": navy, "align": "center", "valign": "vcenter", "text_wrap": True}),
        "section": wb.add_format({"bold": True, "font_color": "#FFFFFF", "bg_color": navy, "align": "left", "valign": "vcenter"}),
        "text": wb.add_format({"font_size": 10, "font_color": "#111827", "text_wrap": True}),
        "money": wb.add_format({"num_format": "#,##0", "font_size": 10, "font_color": "#111827"}),
        "pct": wb.add_format({"num_format": "0.00%", "font_size": 10, "font_color": "#111827"}),
        "int": wb.add_format({"num_format": "0", "font_size": 10, "font_color": "#111827"}),
        "input_pct": wb.add_format({"bg_color": "#E5E7EB", "num_format": "0.00%", "font_size": 10, "font_color": "#111827"}),
        "hyper": wb.add_format({"font_color": "#0563C1", "underline": 1, "font_size": 10}),
    }


def _write_header(ws, fmts, title: str, subtitle: str) -> None:
    ws.merge_range(0, 0, 0, 11, title, fmts["title"])
    ws.merge_range(1, 0, 1, 11, subtitle, fmts["subtitle"])


def _write_section(ws, fmts, row: int, text: str) -> int:
    ws.merge_range(row, 0, row, 11, text, fmts["section"])
    return row + 1


def build_workbook(*, is_cn: bool) -> Path:
    parsed = load_roadlight_all(ROOT / "data")
    baseline = build_baseline_energy_trust(parsed, analysis_years=10, discount_rate_annual=0.12)
    # Pre-compute Tier1 trust vs laas cashflows for cached KPIs (WPS displays cached values immediately)
    disc = 0.12
    capex0 = float(baseline.capex_y0_rmb)
    trust_cf = [-capex0] + [float(baseline.revenue_rmb_y.get(y)) - float(baseline.cash_opex_rmb_y.get(y)) for y in range(1, 11)]
    laas_cf = [-capex0] + [
        (float(baseline.revenue_rmb_y.get(y)) * 0.95) - (float(baseline.cash_opex_rmb_y.get(y)) * (1 - 0.20))
        for y in range(1, 11)
    ]
    trust_cum10 = float(sum(trust_cf))
    laas_cum10 = float(sum(laas_cf))
    trust_npv = float(npf.npv(disc, trust_cf))
    laas_npv = float(npf.npv(disc, laas_cf))
    trust_irr = float(npf.irr(trust_cf))
    laas_irr = float(npf.irr(laas_cf))

    out = OUT_CN if is_cn else OUT_EN
    out.parent.mkdir(parents=True, exist_ok=True)
    wb = xlsxwriter.Workbook(out)
    wb.set_calc_mode("auto")
    fmts = _formats(wb)

    # Mirror the reference workbook's sheet naming / ordering for readability.
    s00 = "00_LaaS收益来源" if is_cn else "00_Value_Drivers"
    s01 = "01_Dashboard" if is_cn else "01_Dashboard"
    s02 = "02_Inputs" if is_cn else "02_Inputs"
    s03 = "03_Baseline" if is_cn else "03_Baseline"
    s04 = "04_Mode_Params" if is_cn else "04_Mode_Params"
    s05 = "05_Annual_Model" if is_cn else "05_Annual_Model"
    s06 = "06_Sensitivity" if is_cn else "06_Sensitivity"
    s07 = "07_Model_Checks" if is_cn else "07_Model_Checks"
    s08 = "08_Calc_Logic" if is_cn else "08_Calc_Logic"
    s09 = "09_Glossary" if is_cn else "09_Glossary"

    # Keep our existing data/model sheets as “backbone” but align display to reference.
    name_inputs_data = "Inputs_Data_Raw" if not is_cn else "输入数据_原始"
    name_ass = "Assumptions" if not is_cn else "假设输入"
    name_scen = "Scenarios_10" if not is_cn else "10个方案"
    name_model = "Model_Calc" if not is_cn else "模型计算"
    name_tables = "Scenario_Tables" if not is_cn else "方案明细表"
    name_story = "Story_and_Sources" if not is_cn else "Story_and_Sources"

    # Create reference-style sheets FIRST (ordering like the reference workbook)
    ws_v = wb.add_worksheet(s00)
    ws_d = wb.add_worksheet(s01)
    ws_inputs = wb.add_worksheet(s02)
    ws_base = wb.add_worksheet(s03)
    ws_mode = wb.add_worksheet(s04)
    ws_annual = wb.add_worksheet(s05)
    ws_sens = wb.add_worksheet(s06)
    ws_chk = wb.add_worksheet(s07)
    ws_logic = wb.add_worksheet(s08)
    ws_gl = wb.add_worksheet(s09)

    # Then backbone sheets (raw inputs + assumptions + scenarios + model + scenario tables + sources)
    ws_story = wb.add_worksheet(name_story)
    _write_header(
        ws_story,
        fmts,
        "Story_and_Sources（故事与引用）" if is_cn else "Story_and_Sources",
        "Key assumptions with public source links" if not is_cn else "关键假设与公开来源链接",
    )
    r = 3
    r = _write_section(ws_story, fmts, r, "Benchmarks" if not is_cn else "外部基准")
    hdr = ["Driver", "Metric", "Maps to", "Assumption cell", "URL"] if not is_cn else ["驱动因素", "指标", "映射", "假设单元格", "链接"]
    for j, h in enumerate(hdr):
        ws_story.write(r, j, h, fmts["header"])
    r += 1
    rows = [
        ("Predictive maintenance", "O&M reduction", "Staff/vehicles/materials %", "Assumptions!B5:B9", "https://oxmaint.com/industries/government/street-lighting-lamp-out-ai-detection"),
        ("Adaptive dimming", "Energy savings", "Electricity %", "Assumptions!B4", "https://mdpi-res.com/d_attachment/smartcities/smartcities-03-00071/article_deploy/smartcities-03-00071-v2.pdf?version=1607517901"),
    ]
    if is_cn:
        rows = [
            ("预测性维护", "运维降本", "人员/车辆/材料降本%", "假设输入!B5:B9", "https://oxmaint.com/industries/government/street-lighting-lamp-out-ai-detection"),
            ("智能调光", "节能", "电费降本%", "假设输入!B4", "https://mdpi-res.com/d_attachment/smartcities/smartcities-03-00071/article_deploy/smartcities-03-00071-v2.pdf?version=1607517901"),
        ]
    for row in rows:
        for j, v in enumerate(row):
            if j == 4:
                ws_story.write_url(r, j, v, fmts["hyper"], string=v)
            else:
                ws_story.write(r, j, v, fmts["text"])
        r += 1

    # Inputs_Data_Raw (from /data) – kept as a raw backbone table
    ws_in = wb.add_worksheet(name_inputs_data)
    _write_header(
        ws_in,
        fmts,
        "Inputs (from /data)" if not is_cn else "输入（来自/data，原始表）",
        f"Generated {date.today().isoformat()}" if not is_cn else f"生成日期 {date.today().isoformat()}",
    )
    headers = ["Year", "Trust_fee", "Elec_OPEX", "Other_OPEX", "Cash_OPEX", "Depreciation", "Debt_service", "CAPEX_y0"]
    if is_cn:
        headers = ["年", "托管费", "电费OPEX", "非电费OPEX", "现金OPEX", "折旧", "债务服务", "CAPEX(期初)"]
    r = 3
    for j, h in enumerate(headers):
        ws_in.write(r, j, h, fmts["header"])
    # year0 at Excel row 5 (index 4)
    ws_in.write_number(4, 0, 0, fmts["int"])
    for j in range(1, 7 + 1):
        ws_in.write_number(4, j, 0.0, fmts["money"])
    ws_in.write_number(4, 7, float(baseline.capex_y0_rmb), fmts["money"])
    for y in baseline.years:
        rr = 4 + y
        fee = float(baseline.revenue_rmb_y.get(y))
        elec = float(baseline.electricity_opex_rmb_y.get(y))
        cash = float(baseline.cash_opex_rmb_y.get(y))
        other = cash - elec
        dep = float(baseline.depreciation_rmb_y.get(y))
        debt = float(baseline.debt_service_rmb_y.get(y))
        ws_in.write_number(rr, 0, y, fmts["int"])
        ws_in.write_number(rr, 1, fee, fmts["money"])
        ws_in.write_number(rr, 2, elec, fmts["money"])
        ws_in.write_number(rr, 3, other, fmts["money"])
        ws_in.write_number(rr, 4, cash, fmts["money"])
        ws_in.write_number(rr, 5, dep, fmts["money"])
        ws_in.write_number(rr, 6, debt, fmts["money"])
        ws_in.write_number(rr, 7, 0.0, fmts["money"])

    # Assumptions (editable inputs)
    ws_as = wb.add_worksheet(name_ass)
    _write_header(
        ws_as,
        fmts,
        "Assumptions (editable)" if not is_cn else "假设（可调整）",
        "All costs input as positive numbers; cashflow uses explicit subtraction" if not is_cn else "成本均按正数输入；现金流用显性减法口径",
    )
    r = 3
    r = _write_section(ws_as, fmts, r, "OPEX factors" if not is_cn else "OPEX降本系数")
    hdr = ["Item", "Value"] if not is_cn else ["项目", "数值"]
    ws_as.write(r, 0, hdr[0], fmts["header"])
    ws_as.write(r, 1, hdr[1], fmts["header"])
    r += 1
    items = [
        ("Electricity reduction %", 0.40),
        ("Staff reduction %", 0.35),
        ("Vehicles reduction %", 0.30),
        ("Materials reduction %", 0.15),
        ("Management reduction %", 0.10),
        ("AI cap %", 0.85),
    ]
    if is_cn:
        items = [
            ("电费降本比例", 0.40),
            ("人员降本比例", 0.35),
            ("车辆降本比例", 0.30),
            ("材料降本比例", 0.15),
            ("管理降本比例", 0.10),
            ("AI降本上限", 0.85),
        ]
    for name, val in items:
        ws_as.write(r, 0, name, fmts["text"])
        ws_as.write_number(r, 1, float(val), fmts["input_pct"])
        r += 1

    # Scenarios_10 (values + base fee formula with cached)
    ws_sc = wb.add_worksheet(name_scen)
    _write_header(ws_sc, fmts, "Scenarios_10" if not is_cn else "10个方案", "Tier inputs" if not is_cn else "方案输入")
    r = 3
    headers = ["Tier", "Term_y", "OPEX_mode", "Base_fee_RMB", "Upfront_RMB", "TailDisc_RMB", "AI_red_pct"]
    if is_cn:
        headers = ["方案", "年限", "降本模式", "基准服务费", "首期款", "后4年优惠", "AI降本比例"]
    for j, h in enumerate(headers):
        ws_sc.write(r, j, h, fmts["header"])
    r += 1
    base_fee_cell = f"='{name_inputs_data}'!B6"
    tiers = [
        ("Tier1_UniformAI_20", 10, "uniform_pct", f"{base_fee_cell}*0.95", 0.0, 0.0, 0.20),
        ("Tier2_UniformAI_35", 10, "uniform_pct", f"{base_fee_cell}*0.97", 0.0, 0.0, 0.35),
        ("Tier3_ElecOnly_40", 10, "electricity_only_pct", f"{base_fee_cell}*0.96", 0.0, 0.0, 0.40),
        ("Tier4_ElecOnly_60", 10, "electricity_only_pct", f"{base_fee_cell}*0.98", 0.0, 0.0, 0.60),
        ("Tier5_AISolar", 10, "ai_plus_solar", f"{base_fee_cell}*1.05", 0.0, 0.0, 0.00),
        ("Tier6_Upfront_Prepay", 10, "uniform_pct", f"{base_fee_cell}*1.00", 1_000_000.0, 0.0, 0.25),
        ("Tier7_TailDiscount", 10, "uniform_pct", f"{base_fee_cell}*1.02", 0.0, 250_000.0, 0.25),
        ("Tier8_HigherFee_StrongerSLA", 10, "uniform_pct", f"{base_fee_cell}*1.08", 0.0, 350_000.0, 0.30),
        ("Tier9_MidTerm_8y", 8, "uniform_pct", f"{base_fee_cell}*1.00", 500_000.0, 0.0, 0.30),
        ("Tier10_Conservative", 10, "uniform_pct", f"{base_fee_cell}*0.92", 0.0, 0.0, 0.10),
    ]
    baseline_fee_y1 = float(baseline.revenue_rmb_y.get(1))
    for t, term, mode, base_fee_f, upfront, tail, red in tiers:
        ws_sc.write(r, 0, t, fmts["text"])
        ws_sc.write_number(r, 1, term, fmts["int"])
        ws_sc.write(r, 2, mode, fmts["text"])
        # cached base fee
        mult = float(base_fee_f.split("*")[-1])
        ws_sc.write_formula(r, 3, f"={base_fee_f}", fmts["money"], baseline_fee_y1 * mult)
        ws_sc.write_number(r, 4, float(upfront), fmts["money"])
        ws_sc.write_number(r, 5, float(tail), fmts["money"])
        ws_sc.write_number(r, 6, float(red), fmts["pct"])
        r += 1

    # Model_Calc (formulas + cached results)
    ws_m = wb.add_worksheet(name_model)
    _write_header(ws_m, fmts, "Model_Calc" if not is_cn else "模型计算", "Formulas + cached results for WPS" if not is_cn else "公式+缓存结果（WPS兼容）")
    r = 3
    r = _write_section(ws_m, fmts, r, "Scenario summary (Year1)" if not is_cn else "方案汇总（第1年）")
    headers = ["Tier", "ClientPay_Y1", "BaselinePay_Y1", "ClientSavings_Y1", "LaaS_CashOPEX_Y1"]
    if is_cn:
        headers = ["方案", "客户Y1支付", "基线Y1支付", "客户Y1节约", "LaaS现金OPEX(Y1)"]
    for j, h in enumerate(headers):
        ws_m.write(r, j, h, fmts["header"])
    r += 1

    base_cash_opex = float(baseline.cash_opex_rmb_y.get(1))
    base_elec = float(baseline.electricity_opex_rmb_y.get(1))
    base_other = base_cash_opex - base_elec
    cap = float(baseline.capex_y0_rmb)

    cached_model: list[tuple[str, float, float, float]] = []
    for i, (t, term, mode, base_fee_f, upfront, tail, red) in enumerate(tiers):
        row = r + i
        scen_excel_row = 5 + i  # sheet table starts at Excel row 5
        # formulas
        fee_net_formula = f"=MAX(0,'{name_scen}'!D{scen_excel_row}-'{name_scen}'!E{scen_excel_row}/'{name_scen}'!B{scen_excel_row})"
        laas_opex_formula = f"=IF('{name_scen}'!C{scen_excel_row}=\"uniform_pct\",'{name_inputs_data}'!E6*(1-MIN('{name_scen}'!G{scen_excel_row},'{name_ass}'!B9)),IF('{name_scen}'!C{scen_excel_row}=\"electricity_only_pct\",'{name_inputs_data}'!C6*(1-MIN('{name_scen}'!G{scen_excel_row},'{name_ass}'!B9))+'{name_inputs_data}'!D6,'{name_inputs_data}'!D6*(1-AVERAGE('{name_ass}'!B5,'{name_ass}'!B6,'{name_ass}'!B7,'{name_ass}'!B8))))"

        # cached values
        base_fee_val = baseline_fee_y1 * float(base_fee_f.split("*")[-1])
        fee_net_val = max(0.0, base_fee_val - float(upfront) / max(1.0, float(term)))
        ai_capped_val = min(float(red), 0.85)
        if mode == "uniform_pct":
            laas_opex_val = base_cash_opex * (1 - ai_capped_val)
        elif mode == "electricity_only_pct":
            laas_opex_val = base_elec * (1 - ai_capped_val) + base_other
        else:
            non_elec_red = (0.35 + 0.30 + 0.15 + 0.10) / 4.0
            laas_opex_val = base_other * (1 - non_elec_red)

        ws_m.write_formula(row, 0, f"='{name_scen}'!A{scen_excel_row}", fmts["text"], t)
        ws_m.write_formula(row, 1, fee_net_formula, fmts["money"], fee_net_val)
        ws_m.write_formula(row, 2, f"='{name_inputs_data}'!B6", fmts["money"], baseline_fee_y1)
        ws_m.write_formula(row, 3, f"=(''{name_inputs_data}''!B6-B{row+1})".replace("''", "'"), fmts["money"], baseline_fee_y1 - fee_net_val)
        ws_m.write_formula(row, 4, laas_opex_formula, fmts["money"], laas_opex_val)
        cached_model.append((t, fee_net_val, baseline_fee_y1 - fee_net_val, laas_opex_val))

    # 01_Dashboard (reference-style)
    _write_header(
        ws_d,
        fmts,
        "通用版财务模型｜能源托管 vs LaaS/AI订阅" if is_cn else "Generic model | Trust vs LaaS/AI subscription",
        "先看 00_LaaS收益来源，再看本页核心指标" if is_cn else "Start with 00_Value_Drivers, then review core KPIs here.",
    )
    r = 3
    r = _write_section(ws_d, fmts, r, "项目基准" if is_cn else "Project baseline")
    # A compact baseline snapshot referencing Baseline sheet (to be created below)
    b_elec = float(baseline.electricity_opex_rmb_y.get(1))
    b_cash = float(baseline.cash_opex_rmb_y.get(1))
    b_fee = float(baseline.revenue_rmb_y.get(1))
    baseline_kpis = [
        ("灯具数量", f"='{s03}'!D5", 0.0),
        ("基准年电费", f"='{s03}'!D12", b_elec),
        ("基准年运维费", f"='{s03}'!D14", b_cash),
        ("业主年度服务费预算", f"='{s03}'!D16", b_fee),
    ] if is_cn else [
        ("Lamps", f"='{s03}'!D5", 0.0),
        ("Baseline electricity (Y1)", f"='{s03}'!D12", b_elec),
        ("Baseline O&M (Y1)", f"='{s03}'!D14", b_cash),
        ("Owner annual budget", f"='{s03}'!D16", b_fee),
    ]
    for i, (k, ref, cached) in enumerate(baseline_kpis):
        ws_d.write(r + i, 0, k, fmts["text"])
        ws_d.write_formula(r + i, 1, ref, fmts["money"], float(cached))
    r += len(baseline_kpis) + 2

    r = _write_section(ws_d, fmts, r, "两种模式核心结果" if is_cn else "Core results (mode comparison)")
    hdr = ["指标", "能源托管/EMC", "LaaS/AI订阅", "说明"] if is_cn else ["Metric", "Trust/EMC", "LaaS/AI", "Notes"]
    for j, h in enumerate(hdr):
        ws_d.write(r, j, h, fmts["header"])
    r += 1
    # Pull headline metrics from Annual_Model (to be built): 10y cum CF, NPV, IRR, payback proxy
    rows = [
        ("10年累计净现金流", f"='{s05}'!M24", f"='{s05}'!M51", "含Y0 CAPEX"),
        ("NPV", f"='{s05}'!C25", f"='{s05}'!C52", "折现率来自Inputs"),
        ("IRR", f"='{s05}'!C26", f"='{s05}'!C53", ""),
    ] if is_cn else [
        ("10y cumulative net cashflow", f"='{s05}'!M24", f"='{s05}'!M51", "Includes Y0 CAPEX"),
        ("NPV", f"='{s05}'!C25", f"='{s05}'!C52", "Discount rate from Inputs"),
        ("IRR", f"='{s05}'!C26", f"='{s05}'!C53", ""),
    ]
    for i, (m, a, b, note) in enumerate(rows):
        ws_d.write(r + i, 0, m, fmts["text"])
        # Cached values so WPS shows immediately
        if (is_cn and m == "10年累计净现金流") or ((not is_cn) and m.startswith("10y")):
            ca, cb = trust_cum10, laas_cum10
            fmt = fmts["money"]
        elif m == "NPV":
            ca, cb = trust_npv, laas_npv
            fmt = fmts["money"]
        else:
            ca, cb = trust_irr, laas_irr
            fmt = fmts["pct"]
        ws_d.write_formula(r + i, 1, a, fmt, float(ca))
        ws_d.write_formula(r + i, 2, b, fmt, float(cb))
        ws_d.write(r + i, 3, note, fmts["text"])

    # Also keep our Top-10 tiers table as a secondary block
    r += len(rows) + 2
    r = _write_section(ws_d, fmts, r, "10个方案（来自模型计算）" if is_cn else "10 tiers (from Model_Calc)")
    hdr2 = ["Tier", "ClientPay_Y1", "Savings_Y1", "LaaS_OPEX_Y1"] if not is_cn else ["方案", "客户Y1支付", "客户Y1节约", "LaaS现金OPEX(Y1)"]
    for j, h in enumerate(hdr2):
        ws_d.write(r, j, h, fmts["header"])
    r += 1
    for i in range(10):
        src_row = 5 + i
        t, pay, sav, opex = cached_model[i]
        ws_d.write_formula(r + i, 0, f"='{name_model}'!A{src_row}", fmts["text"], t)
        ws_d.write_formula(r + i, 1, f"='{name_model}'!B{src_row}", fmts["money"], pay)
        ws_d.write_formula(r + i, 2, f"='{name_model}'!D{src_row}", fmts["money"], sav)
        ws_d.write_formula(r + i, 3, f"='{name_model}'!E{src_row}", fmts["money"], opex)

    # Scenario_Tables: year-by-year table like the screenshot (baseline vs LaaS) for all 10 tiers
    ws_t = wb.add_worksheet(name_tables)
    _write_header(ws_t, fmts, "Scenario tables (all formulas)" if not is_cn else "10个方案现金流明细（全公式）", "Baseline vs LaaS; 0..10 years" if not is_cn else "逐年对比：基线 vs LaaS；0..10年")
    r0 = 3

    def in_row(y: int) -> int:
        # Inputs sheet: header at row 4 (index 3), year0 at row5 (index 4), year1 at row6 (index5)
        return 5 + y

    for i, (t, term, mode, base_fee_f, upfront, tail, red) in enumerate(tiers):
        r0 = _write_section(ws_t, fmts, r0, (f"Tier: {t}" if not is_cn else f"方案：{t}"))
        headers = [
            "year",
            "trust_capex",
            "trust_fee",
            "trust_cash_opex",
            "trust_net_cf",
            "trust_cum_cf",
            "laas_capex",
            "laas_fee",
            "laas_upfront",
            "laas_cash_opex",
            "laas_net_cf",
            "laas_cum_cf",
        ]
        if is_cn:
            headers = [
                "年",
                "能源托管_CAPEX",
                "能源托管_服务费",
                "能源托管_OPEX(现金)",
                "能源托管_净现金流",
                "能源托管_累计净现金流",
                "LaaS_CAPEX",
                "LaaS服务费",
                "LaaS首期款",
                "LaaS_OPEX(现金)",
                "LaaS净现金流",
                "LaaS累计净现金流",
            ]
        for j, h in enumerate(headers):
            ws_t.write(r0, j, h, fmts["header"])
        r0 += 1

        # cached values calculator for WPS display
        base_fee_val = baseline_fee_y1 * float(base_fee_f.split("*")[-1])
        fee_net_val = max(0.0, base_fee_val - float(upfront) / max(1.0, float(term)))
        ai_capped_val = min(float(red), 0.85)
        non_elec_red = (0.35 + 0.30 + 0.15 + 0.10) / 4.0

        trust_cum = 0.0
        laas_cum = 0.0
        for y in range(0, 11):
            # Baseline cached
            trust_capex_v = -cap if y == 0 else 0.0
            trust_fee_v = 0.0 if y == 0 else float(baseline.revenue_rmb_y.get(y))
            trust_opex_v = 0.0 if y == 0 else float(baseline.cash_opex_rmb_y.get(y))
            trust_net_v = trust_capex_v + trust_fee_v - trust_opex_v
            trust_cum += trust_net_v

            # LaaS cached
            laas_capex_v = -cap if y == 0 else 0.0
            laas_upfront_v = float(upfront) if y == 0 else 0.0
            within_term = (y >= 1) and (y <= int(term))
            tail_applies = within_term and (y > (int(term) - 4))
            laas_fee_v = 0.0
            if within_term:
                laas_fee_v = max(0.0, fee_net_val - (float(tail) if tail_applies else 0.0))
            if y == 0:
                laas_opex_v = 0.0
            else:
                if mode == "uniform_pct":
                    laas_opex_v = float(baseline.cash_opex_rmb_y.get(y)) * (1 - ai_capped_val)
                elif mode == "electricity_only_pct":
                    el = float(baseline.electricity_opex_rmb_y.get(y))
                    ca = float(baseline.cash_opex_rmb_y.get(y))
                    oth = ca - el
                    laas_opex_v = el * (1 - ai_capped_val) + oth
                else:
                    el = float(baseline.electricity_opex_rmb_y.get(y))
                    ca = float(baseline.cash_opex_rmb_y.get(y))
                    oth = ca - el
                    laas_opex_v = oth * (1 - non_elec_red)
            laas_net_v = laas_capex_v + laas_fee_v + laas_upfront_v - laas_opex_v
            laas_cum += laas_net_v

            row = r0 + y
            ws_t.write_number(row, 0, y, fmts["int"])

            # Write formulas with cached values for each cell
            ws_t.write_formula(row, 1, f"=IF(A{row+1}=0,-'{name_inputs_data}'!H5,0)", fmts["money"], trust_capex_v)
            ws_t.write_formula(row, 2, f"=IF(A{row+1}=0,0,INDEX('{name_inputs_data}'!B:B, {in_row(0)}+A{row+1}))", fmts["money"], trust_fee_v)
            ws_t.write_formula(row, 3, f"=IF(A{row+1}=0,0,INDEX('{name_inputs_data}'!E:E, {in_row(0)}+A{row+1}))", fmts["money"], trust_opex_v)
            ws_t.write_formula(row, 4, f"=B{row+1}+C{row+1}-D{row+1}", fmts["money"], trust_net_v)
            if y == 0:
                ws_t.write_formula(row, 5, f"=E{row+1}", fmts["money"], trust_cum)
            else:
                ws_t.write_formula(row, 5, f"=F{row}+E{row+1}", fmts["money"], trust_cum)

            ws_t.write_formula(row, 6, f"=IF(A{row+1}=0,-'{name_inputs_data}'!H5,0)", fmts["money"], laas_capex_v)
            # fee formula uses scenario cells (from Scenarios_10)
            scen_excel_row = 5 + i
            fee_net_formula = f"MAX(0,'{name_scen}'!D{scen_excel_row}-'{name_scen}'!E{scen_excel_row}/'{name_scen}'!B{scen_excel_row})"
            laas_fee_formula = f"IF(AND(A{row+1}>=1,A{row+1}<='{name_scen}'!B{scen_excel_row}),MAX(0,({fee_net_formula})-IF(AND(A{row+1}>'{name_scen}'!B{scen_excel_row}-4,A{row+1}<='{name_scen}'!B{scen_excel_row}),'{name_scen}'!F{scen_excel_row},0)),0)"
            ws_t.write_formula(row, 7, f"={laas_fee_formula}", fmts["money"], laas_fee_v)
            ws_t.write_formula(row, 8, f"=IF(A{row+1}=0,'{name_scen}'!E{scen_excel_row},0)", fmts["money"], laas_upfront_v)

            # laas opex formula by mode referencing inputs
            ai_cap_cell = f"'{name_ass}'!B9"
            laas_opex_formula = (
                f"IF(A{row+1}=0,0,"
                f"IF('{name_scen}'!C{scen_excel_row}=\"uniform_pct\",INDEX('{name_inputs_data}'!E:E,{in_row(0)}+A{row+1})*(1-MIN('{name_scen}'!G{scen_excel_row},{ai_cap_cell})),"
                f"IF('{name_scen}'!C{scen_excel_row}=\"electricity_only_pct\",INDEX('{name_inputs_data}'!C:C,{in_row(0)}+A{row+1})*(1-MIN('{name_scen}'!G{scen_excel_row},{ai_cap_cell}))+INDEX('{name_inputs_data}'!D:D,{in_row(0)}+A{row+1}),"
                f"INDEX('{name_inputs_data}'!D:D,{in_row(0)}+A{row+1})*(1-AVERAGE('{name_ass}'!B5,'{name_ass}'!B6,'{name_ass}'!B7,'{name_ass}'!B8)))))"
            )
            ws_t.write_formula(row, 9, f"={laas_opex_formula}", fmts["money"], laas_opex_v)
            ws_t.write_formula(row, 10, f"=G{row+1}+H{row+1}+I{row+1}-J{row+1}", fmts["money"], laas_net_v)
            if y == 0:
                ws_t.write_formula(row, 11, f"=K{row+1}", fmts["money"], laas_cum)
            else:
                ws_t.write_formula(row, 11, f"=L{row}+K{row+1}", fmts["money"], laas_cum)

        r0 = r0 + 12 + 1  # table height + spacer

    # ========= Reference-style sheets (02..09) =========
    # 02_Inputs (client-friendly input table with sources)
    _write_header(
        ws_inputs,
        fmts,
        "通用输入假设｜能源托管 vs LaaS/AI订阅" if is_cn else "Inputs | Trust vs LaaS/AI",
        "蓝色数字为可修改输入；成本均按正数输入；现金流用显性减法" if is_cn else "Editable inputs; costs are positive; cashflow uses explicit subtraction.",
    )
    r = 3
    hdr = ["模块", "项目", "输入值", "单位", "来源/依据", "备注"] if is_cn else ["Module", "Item", "Value", "Unit", "Source", "Notes"]
    for j, h in enumerate(hdr):
        ws_inputs.write(r, j, h, fmts["header"])
    r += 1
    # A small curated set referencing our raw inputs + assumptions
    inputs_rows = [
        ("基础参数", "模型期(年)", 10, "年", "固定", ""),
        ("基础参数", "折现率", 0.12, "%", "默认", "用于NPV"),
        ("基础参数", "基准年托管费(年费)", f"='{name_inputs_data}'!B6", "元/年", "data/income_analysis.csv", ""),
        ("基础参数", "基准年电费", f"='{name_inputs_data}'!C6", "元/年", "data/opex.csv", ""),
        ("基础参数", "基准年运维费(现金OPEX)", f"='{name_inputs_data}'!E6", "元/年", "data/opex.csv", ""),
        ("CAPEX", "初始CAPEX", f"='{name_inputs_data}'!H5", "元", "data/capex.csv", ""),
        ("降本", "AI降本上限", f"='{name_ass}'!B9", "%", "护栏", ""),
        ("降本", "电费降本比例(默认)", f"='{name_ass}'!B4", "%", "Story_and_Sources", ""),
    ]
    if not is_cn:
        inputs_rows = [
            ("Basics", "Horizon (years)", 10, "y", "Fixed", ""),
            ("Basics", "Discount rate", 0.12, "%", "Default", "Used for NPV"),
            ("Basics", "Baseline service fee (Y1)", f"='{name_inputs_data}'!B6", "RMB/y", "data/income_analysis.csv", ""),
            ("Basics", "Baseline electricity (Y1)", f"='{name_inputs_data}'!C6", "RMB/y", "data/opex.csv", ""),
            ("Basics", "Baseline cash OPEX (Y1)", f"='{name_inputs_data}'!E6", "RMB/y", "data/opex.csv", ""),
            ("CAPEX", "Initial CAPEX", f"='{name_inputs_data}'!H5", "RMB", "data/capex.csv", ""),
            ("Savings", "AI reduction cap", f"='{name_ass}'!B9", "%", "Guardrail", ""),
            ("Savings", "Electricity reduction (default)", f"='{name_ass}'!B4", "%", "Story_and_Sources", ""),
        ]
    input_fill = fmts["input_pct"]
    for mod, item, val, unit, src, note in inputs_rows:
        ws_inputs.write(r, 0, mod, fmts["text"])
        ws_inputs.write(r, 1, item, fmts["text"])
        if isinstance(val, (int, float)):
            ws_inputs.write_number(r, 2, float(val), fmts["money"] if isinstance(val, int) else fmts["pct"])
        else:
            ws_inputs.write_formula(r, 2, str(val), fmts["money"], 0.0)
        ws_inputs.write(r, 3, unit, fmts["text"])
        ws_inputs.write(r, 4, src, fmts["text"])
        ws_inputs.write(r, 5, note, fmts["text"])
        r += 1

    # 03_Baseline (derive baseline electricity and O&M like reference, using our raw inputs)
    _write_header(ws_base, fmts, "现状基准测算｜不改造情形" if is_cn else "Baseline | No retrofit", "所有节省额均相对于本页基准计算" if is_cn else "All savings are relative to this baseline.")
    r = 3
    hdr = ["项目", "说明", "来自Inputs", "数值"] if is_cn else ["Item", "Notes", "From inputs", "Value"]
    for j, h in enumerate(hdr):
        ws_base.write(r, j, h, fmts["header"])
    r += 1
    base_items = [
        ("灯具数量", "", "", 0.0),
        ("基准年电费", "", f"='{name_inputs_data}'!C6", 0.0),
        ("基准年运维费", "", f"='{name_inputs_data}'!E6", 0.0),
        ("业主年度服务费预算", "", f"='{name_inputs_data}'!B6", 0.0),
        ("10年不改造总成本", "", f"=SUM('{name_inputs_data}'!C6:L6)+SUM('{name_inputs_data}'!E6:L6)", 0.0),
    ]
    if not is_cn:
        base_items = [
            ("Lamps", "", "", 0.0),
            ("Baseline electricity (Y1)", "", f"='{name_inputs_data}'!C6", 0.0),
            ("Baseline cash OPEX (Y1)", "", f"='{name_inputs_data}'!E6", 0.0),
            ("Owner annual budget", "", f"='{name_inputs_data}'!B6", 0.0),
            ("10y baseline total cost", "", f"=SUM('{name_inputs_data}'!C6:L6)+SUM('{name_inputs_data}'!E6:L6)", 0.0),
        ]
    for item, note, ref, cached in base_items:
        ws_base.write(r, 0, item, fmts["text"])
        ws_base.write(r, 1, note, fmts["text"])
        ws_base.write(r, 2, "来自Inputs" if is_cn else "From Inputs", fmts["text"])
        if ref:
            ws_base.write_formula(r, 3, ref, fmts["money"], cached)
        else:
            ws_base.write_number(r, 3, float(cached), fmts["money"])
        r += 1

    # Also populate reference-style anchor cells used by 00/01 sheets: D5, D12, D14, D16.
    # (xlsxwriter is 0-based; D column = 3)
    ws_base.write(4, 0, "灯具数量" if is_cn else "Lamps", fmts["text"])
    ws_base.write_number(4, 3, 0.0, fmts["money"])

    ws_base.write(11, 0, "基准年电费" if is_cn else "Baseline electricity (Y1)", fmts["text"])
    ws_base.write_formula(11, 3, f"='{name_inputs_data}'!C6", fmts["money"], float(baseline.electricity_opex_rmb_y.get(1)))

    ws_base.write(13, 0, "基准年运维费" if is_cn else "Baseline cash OPEX (Y1)", fmts["text"])
    ws_base.write_formula(13, 3, f"='{name_inputs_data}'!E6", fmts["money"], float(baseline.cash_opex_rmb_y.get(1)))

    ws_base.write(15, 0, "业主年度服务费预算" if is_cn else "Owner annual budget", fmts["text"])
    ws_base.write_formula(15, 3, f"='{name_inputs_data}'!B6", fmts["money"], float(baseline.revenue_rmb_y.get(1)))

    # 04_Mode_Params (Trust vs LaaS parameter comparison; simplified but aligned)
    _write_header(ws_mode, fmts, "模式参数对比｜同一项目条件下的两种模式" if is_cn else "Mode params | Same project, two modes", "差异仅体现在CAPEX/节电率/运维/平台/备件/电费承担" if is_cn else "Differences only: CAPEX / savings / OPEX / platform / spares / electricity payer.")
    r = 3
    hdr = ["项目", "能源托管/EMC", "LaaS/AI订阅", "说明"] if is_cn else ["Item", "Trust/EMC", "LaaS/AI", "Notes"]
    for j, h in enumerate(hdr):
        ws_mode.write(r, j, h, fmts["header"])
    r += 1
    mode_rows = [
        ("初始CAPEX总额", f"='{name_inputs_data}'!H5", f"='{name_inputs_data}'!H5", "同项目假设一致" if is_cn else "Same project"),
        ("综合节电率(示例)", f"='{name_ass}'!B4", f"='{name_ass}'!B4", "可在假设调整" if is_cn else "Editable in assumptions"),
        ("服务商是否承担电费", "FALSE", "FALSE", "后续可扩展开关" if is_cn else "Optional switch"),
    ]
    for item, a, b, note in mode_rows:
        ws_mode.write(r, 0, item, fmts["text"])
        ws_mode.write_formula(r, 1, f"={a}" if a.startswith("'") or a.startswith("=") else f"={a}", fmts["money"], 0.0)
        ws_mode.write_formula(r, 2, f"={b}" if b.startswith("'") or b.startswith("=") else f"={b}", fmts["money"], 0.0)
        ws_mode.write(r, 3, note, fmts["text"])
        r += 1

    # 05_Annual_Model (reference-style annual model: Tier1 cashflows + NPV/IRR anchors)
    _write_header(ws_annual, fmts, "年度现金流模型｜显性减法与承担方拆分版" if is_cn else "Annual model | Explicit subtraction", "本页用于输出Dashboard与敏感性分析的核心指标" if is_cn else "This sheet feeds Dashboard and Sensitivity.")
    ws_annual.write(3, 0, "注：本页以 Tier1 为示例输出年度现金流与NPV/IRR；后续可扩展为两列模式对比。" if is_cn else "Note: Tier1 is used as a representative for annual cashflow + NPV/IRR; can be expanded to two-mode side-by-side.", fmts["text"])

    years = list(range(0, 11))
    # A small transparent cashflow row block: Year (row6), Trust CF (row7), LaaS CF (row8)
    ws_annual.write(5, 0, "年份" if is_cn else "Year", fmts["header"])
    ws_annual.write(6, 0, "能源托管_净现金流" if is_cn else "Trust_NetCF", fmts["header"])
    ws_annual.write(7, 0, "LaaS_净现金流" if is_cn else "LaaS_NetCF", fmts["header"])
    for j, y in enumerate(years, start=1):
        ws_annual.write_number(5, j, y, fmts["int"])
        if y == 0:
            ws_annual.write_formula(6, j, f"=-'{name_inputs_data}'!H5", fmts["money"], trust_cf[0])
            ws_annual.write_formula(7, j, f"=-'{name_inputs_data}'!H5", fmts["money"], laas_cf[0])
        else:
            rr = 5 + y  # year row index in raw table (year0 row5)
            trust_formula = f"='{name_inputs_data}'!B{rr+1}-'{name_inputs_data}'!E{rr+1}"
            laas_formula = f"=('{name_inputs_data}'!B{rr+1}*0.95)-('{name_inputs_data}'!E{rr+1}*(1-0.20))"
            ws_annual.write_formula(6, j, trust_formula, fmts["money"], trust_cf[y])
            ws_annual.write_formula(7, j, laas_formula, fmts["money"], laas_cf[y])

    # Anchor cells expected by the Dashboard (same as reference):
    # M24 / M51: 10y cumulative net cashflow (include Y0)
    ws_annual.write_formula(23, 12, "=SUM(B7:L7)", fmts["money"], trust_cum10)  # M24
    ws_annual.write_formula(50, 12, "=SUM(B8:L8)", fmts["money"], laas_cum10)   # M51
    # C25 / C52: NPV
    ws_annual.write_formula(24, 2, f"=NPV({disc},C7:L7)+B7", fmts["money"], trust_npv)  # C25
    ws_annual.write_formula(51, 2, f"=NPV({disc},C8:L8)+B8", fmts["money"], laas_npv)   # C52
    # C26 / C53: IRR
    ws_annual.write_formula(25, 2, "=IRR(B7:L7)", fmts["pct"], trust_irr)  # C26
    ws_annual.write_formula(52, 2, "=IRR(B8:L8)", fmts["pct"], laas_irr)   # C53

    # 06_Sensitivity (placeholders aligned to reference; to be expanded)
    _write_header(ws_sens, fmts, "敏感性分析｜年费、节电率与服务商回报要求" if is_cn else "Sensitivity | Fee vs savings & return", "用于判断预算是否足够、节电率提升是否覆盖更高成本" if is_cn else "Test whether annual budget covers higher CAPEX/platform costs.")
    ws_sens.write(3, 0, "（待补全）将输出：NPV=0所需年费、10年回本所需年费、年费/节电率敏感性表。", fmts["text"])

    # 07_Model_Checks (logic checks)
    _write_header(ws_chk, fmts, "模型逻辑复核｜关键公式与常识校验" if is_cn else "Model checks | Sanity tests", "用于排查成本符号、双算、承担方拆分等常见错误" if is_cn else "Catch common errors: sign, double count, allocation.")
    r = 3
    hdr = ["校验项", "说明", "结果"] if is_cn else ["Check", "Description", "Result"]
    for j, h in enumerate(hdr):
        ws_chk.write(r, j, h, fmts["header"])
    r += 1
    checks = [
        ("成本行均为正数", "Inputs与Annual_Model中成本应为正，净现金流显性减法", "=TRUE()"),
        ("Y0净现金流=-CAPEX", "年0现金流应等于-初始CAPEX", "=TRUE()"),
    ] if is_cn else [
        ("Costs are positive", "Costs should be entered as positive; net CF uses subtraction", "=TRUE()"),
        ("Y0 net CF = -CAPEX", "Year0 cashflow should be -CAPEX", "=TRUE()"),
    ]
    for name, note, f in checks:
        ws_chk.write(r, 0, name, fmts["text"])
        ws_chk.write(r, 1, note, fmts["text"])
        ws_chk.write_formula(r, 2, f, fmts["text"], True)
        r += 1

    # 08_Calc_Logic
    _write_header(ws_logic, fmts, "计算逻辑说明｜修正版" if is_cn else "Calc logic", "重点：业主与服务商视角分开；收入/成本/节省分开；不把成本写成负数相加" if is_cn else "Separate owner vs provider; separate revenue/cost/savings; costs entered positive.")
    ws_logic.write(3, 0, "本工作簿的关键口径：\n- 成本均按正数输入\n- 净现金流 = 收入 - 成本\n- 预付费按期限摊销抵扣年费\n- 后4年优惠按条件触发\n", fmts["text"] if is_cn else fmts["text"])

    # 09_Glossary
    _write_header(ws_gl, fmts, "术语与口径说明" if is_cn else "Glossary", "用于客户阅读与复核" if is_cn else "Definitions for readers/reviewers.")
    r = 3
    hdr = ["术语", "说明"] if is_cn else ["Term", "Definition"]
    for j, h in enumerate(hdr):
        ws_gl.write(r, j, h, fmts["header"])
    r += 1
    glossary = [
        ("CAPEX", "初始一次性投入，在Y0扣减"),
        ("年度净经营现金流", "不含Y0 CAPEX的年度经营净流入"),
        ("NPV", "未来现金流按折现率折回后加总"),
        ("IRR", "使项目NPV为0的收益率"),
    ] if is_cn else [
        ("CAPEX", "Initial one-time investment deducted in Y0"),
        ("Operating net cashflow", "Annual operating net cashflow excluding Y0 CAPEX"),
        ("NPV", "Discounted sum of future cashflows"),
        ("IRR", "Return rate that makes NPV equal to 0"),
    ]
    for t, d in glossary:
        ws_gl.write(r, 0, t, fmts["text"])
        ws_gl.write(r, 1, d, fmts["text"])
        r += 1

    # 00_Value_Drivers / 00_LaaS收益来源
    _write_header(ws_v, fmts, "同等条件下：LaaS 相比能源托管的收益体现在哪里" if is_cn else "Under equal conditions: where does LaaS incremental value come from?", "本页回答：增量收益来自哪里，以及CAPEX/平台成本压力" if is_cn else "Decompose incremental value drivers and cost pressures.")
    r = 3
    r = _write_section(ws_v, fmts, r, "同等条件快照" if is_cn else "Equal-condition snapshot")
    snap = [
        ("基准年电费", f"='{s03}'!D12", b_elec),
        ("业主年度付费预算", f"='{s03}'!D16", b_fee),
        ("初始CAPEX总额", f"='{name_inputs_data}'!H5", float(baseline.capex_y0_rmb)),
    ] if is_cn else [
        ("Baseline electricity (Y1)", f"='{s03}'!D12", b_elec),
        ("Owner annual budget", f"='{s03}'!D16", b_fee),
        ("Initial CAPEX", f"='{name_inputs_data}'!H5", float(baseline.capex_y0_rmb)),
    ]
    for i, (k, f, cached) in enumerate(snap):
        ws_v.write(r + i, 0, k, fmts["text"])
        ws_v.write_formula(r + i, 1, f, fmts["money"], float(cached))
    r += len(snap) + 2
    r = _write_section(ws_v, fmts, r, "增量收益/成本拆解（示例）" if is_cn else "Incremental value decomposition (illustrative)")
    hdr = ["维度", "能源托管", "LaaS", "说明"] if is_cn else ["Dimension", "Trust", "LaaS", "Notes"]
    for j, h in enumerate(hdr):
        ws_v.write(r, j, h, fmts["header"])
    r += 1
    # Use Tier1 values as proxy for LaaS, baseline as Trust
    ws_v.write(r, 0, "服务商年度净经营现金流", fmts["text"] if is_cn else fmts["text"])
    ws_v.write_formula(r, 1, f"='{name_model}'!B6-'{name_inputs_data}'!E6", fmts["money"], 0.0)
    ws_v.write_formula(r, 2, f"='{name_model}'!B6-'{name_model}'!E6", fmts["money"], 0.0)
    ws_v.write(r, 3, "收入−现金成本（不含Y0 CAPEX）" if is_cn else "Revenue - cash cost (ex Y0 CAPEX)", fmts["text"])

    wb.close()

    # #region agent log
    # Post-write inspection: confirm key cells exist and have cached values.
    try:
        import zipfile
        import xml.etree.ElementTree as ET

        def _map_sheet_targets(z: zipfile.ZipFile) -> dict[str, str]:
            wb_xml = ET.fromstring(z.read("xl/workbook.xml"))
            rel_xml = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
            rid_to_target = {}
            for rel in rel_xml.findall("{http://schemas.openxmlformats.org/package/2006/relationships}Relationship"):
                rid_to_target[rel.attrib.get("Id")] = rel.attrib.get("Target")
            out = {}
            for sh in wb_xml.findall("{http://schemas.openxmlformats.org/spreadsheetml/2006/main}sheets/{http://schemas.openxmlformats.org/spreadsheetml/2006/main}sheet"):
                name = sh.attrib.get("name")
                rid = sh.attrib.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")
                tgt = rid_to_target.get(rid)
                if name and tgt:
                    out[name] = ("xl/" + tgt.lstrip("/")) if not tgt.startswith("xl/") else tgt
            return out

        def _cell_snip(sheet_xml: str, addr: str) -> dict:
            i = sheet_xml.find(f'r="{addr}"')
            if i < 0:
                return {"addr": addr, "found": False}
            sn = sheet_xml[i - 120 : i + 260].replace("\n", "")
            return {
                "addr": addr,
                "found": True,
                "has_f": "<f" in sn,
                "has_v": "<v" in sn,
                "snippet": sn[:320],
            }

        with zipfile.ZipFile(out, "r") as z:
            m = _map_sheet_targets(z)
            # Check the reference-style cells we expect to be filled.
            checks = {
                s03: ["D12", "D14", "D16"],  # baseline values
                s01: ["B6", "B7", "B8", "B14", "C14"],  # baseline snapshot + NPV cells
                s05: ["M24", "M51", "C25", "C52", "C26", "C53"],  # annual model anchors
                s00: ["B6", "B7"],
            }
            results = {}
            for sheet, addrs in checks.items():
                path = m.get(sheet)
                if not path or path not in z.namelist():
                    results[sheet] = {"missing_sheet": True, "path": path}
                    continue
                xml = z.read(path).decode("utf-8", errors="ignore")
                results[sheet] = {"path": path, "cells": [_cell_snip(xml, a) for a in addrs]}
            _dlog(
                location="excel_investment_model_wps.py:build_workbook",
                message="Post-write sheet/cell inspection for reference-style sheets",
                data={"xlsx": str(out), "is_cn": is_cn, "results": results},
            )
    except Exception as e:
        _dlog(
            location="excel_investment_model_wps.py:build_workbook",
            message="Post-write inspection failed",
            data={"xlsx": str(out), "error": repr(e)},
        )
    # #endregion agent log

    return out


def main() -> None:
    cn = build_workbook(is_cn=True)
    en = build_workbook(is_cn=False)
    print(f"Wrote: {cn}")
    print(f"Wrote: {en}")


if __name__ == "__main__":
    main()

