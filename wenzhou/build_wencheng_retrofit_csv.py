"""Generate wenzhou/wencheng_retrofit_and_om_baseline.csv (retrofit table + O&M benchmark)."""

from __future__ import annotations

import csv
from pathlib import Path

ROOT = Path(__file__).resolve().parent
OUT = ROOT / "wencheng_retrofit_and_om_baseline.csv"

HDR = [
    "record_type",
    "row_id",
    "luminaire_type_zh",
    "luminaire_type_en",
    "power_w",
    "qty_units",
    "hours_per_year",
    "annual_kwh_theoretical",
    "delta_power_w",
    "annual_kwh_saved_line",
    "tariff_cny_per_kwh",
    "annual_fee_cny",
    "unit_capex_cny",
    "line_capex_cny",
    "ratio_name",
    "ratio_value",
    "currency",
    "om_component_zh",
    "om_cny_per_lamp_year",
    "om_annual_total_cny",
    "notes",
]


def cells(**kw: object) -> list[str]:
    """Emit one CSV row; unspecified HDR columns become empty string."""
    rowd: dict[str, str] = {h: "" for h in HDR}
    for k, v in kw.items():
        if k not in HDR:
            raise KeyError(k)
        if v is None or v == "":
            rowd[k] = ""
        elif isinstance(v, float):
            rowd[k] = str(v)
        else:
            rowd[k] = str(v)
    return [rowd[h] for h in HDR]


def main() -> None:
    rows: list[list[str]] = []

    rows.append(
        cells(
            record_type="meta",
            row_id="meta_01",
            luminaire_type_zh="文成隧道路灯改造测算",
            currency="CNY",
            notes=(
                "source=user spreadsheet; tariff=0.72 CNY/kWh; "
                "survey_column_retrofit_after_electricity_fee_cny = sum(savings_line annual_fee_cny) = saved_kWh_total * λ (~483636); "
                "annual_savings_fee_cny = footer_before_kWh*λ - that = footer_after_kWh*λ (~412618); "
                "retrofit_after kWh lines = before_kWh minus saved_kWh per row; "
                "stated_footer_after_kWh = footer_before_kWh - sum(saved_kWh); "
                "reconcile when line-sum before != footer before (same as agg_before)."
            ),
        )
    )

    before_lines = [
        ("before_01", "专用公路LED隧道灯", "LED tunnel (highway special)", 90, 654, 8760, 515652),
        ("before_02", "专用公路LED隧道灯", "LED tunnel (highway special)", 100, 182, 8760, 159432),
        ("before_03", "专用公路LED隧道灯", "LED tunnel (highway special)", 150, 282, 8760, 370548),
        ("before_04", "专用公路LED隧道灯", "LED tunnel (highway special)", 250, 26, 8760, 56940),
        ("before_05", "LED路灯", "LED street", 150, 532, 4380, 415221),
    ]
    for bid, zht, ent, pw, q, h, ak in before_lines:
        rows.append(
            cells(
                record_type="retrofit_before",
                row_id=bid,
                luminaire_type_zh=zht,
                luminaire_type_en=ent,
                power_w=pw,
                qty_units=q,
                hours_per_year=h,
                annual_kwh_theoretical=ak,
                currency="CNY",
            )
        )

    sum_before_lines = sum(x[-1] for x in before_lines)
    footer_before = 1244798

    rows.append(
        cells(
            record_type="aggregate",
            row_id="agg_before_kwh_sum_lines",
            ratio_name="sum_of_line_items_before_kwh",
            ratio_value=sum_before_lines,
            currency="CNY",
        )
    )
    rows.append(
        cells(
            record_type="aggregate",
            row_id="agg_before_kwh_footer",
            ratio_name="stated_footer_before_kwh",
            ratio_value=footer_before,
            currency="CNY",
        )
    )
    note_rec = (
        ""
        if abs(sum_before_lines - footer_before) < 0.01
        else (
            "reconcile_required: sum_of_line_items_kWh ("
            + str(sum_before_lines)
            + ") != stated_footer ("
            + str(footer_before)
            + ")"
        )
    )
    rows.append(
        cells(
            record_type="aggregate",
            row_id="agg_before_kwh_reconcile_note",
            ratio_name="line_minus_footer_kwh_delta",
            ratio_value=sum_before_lines - footer_before,
            currency="CNY",
            notes=note_rec,
        )
    )

    # 节电量（kWh/年）：收资表「节能量块」分项（表头可能写作年理论电量，实为逐年节电量）；与 savings_line 一致。
    saved_kwh_by_row = [145591.2, 79716.0, 197625.6, 27331.2, 221452.8]
    total_saved_kwh = float(sum(saved_kwh_by_row))

    after_specs = [
        ("after_01", "LED隧道灯", "LED tunnel", 40, 654, 8760),
        ("after_02", "LED隧道灯", "LED tunnel", 50, 182, 8760),
        ("after_03", "LED隧道灯", "LED tunnel", 100, 282, 8760),
        ("after_04", "LED隧道灯", "LED tunnel", 200, 26, 8760),
        ("after_05", "LED路灯", "LED street", 120, 532, 4380),
    ]
    after_lines: list[tuple[str, str, str, int, int, int, float]] = []
    for i, (aid, zht, ent, pw, q, h) in enumerate(after_specs):
        ak_before_row = float(before_lines[i][-1])
        ak_after = ak_before_row - saved_kwh_by_row[i]
        after_lines.append((aid, zht, ent, pw, q, h, ak_after))

    for aid, zht, ent, pw, q, h, ak in after_lines:
        rows.append(
            cells(
                record_type="retrofit_after",
                row_id=aid,
                luminaire_type_zh=zht,
                luminaire_type_en=ent,
                power_w=pw,
                qty_units=q,
                hours_per_year=h,
                annual_kwh_theoretical=ak,
                currency="CNY",
            )
        )

    sum_after_lines = sum(x[-1] for x in after_lines)
    # 表尾：改造后 = 表尾改造前 − 节电量合计（与分项加总「改造后」可能不一致，当「改造前」表尾≠分项加总时）
    footer_after = float(footer_before) - total_saved_kwh

    rows.append(
        cells(
            record_type="aggregate",
            row_id="agg_after_kwh_sum_lines",
            ratio_name="sum_of_line_items_after_kwh",
            ratio_value=sum_after_lines,
            currency="CNY",
        )
    )
    rows.append(
        cells(
            record_type="aggregate",
            row_id="agg_after_kwh_footer",
            ratio_name="stated_footer_after_kwh",
            ratio_value=footer_after,
            currency="CNY",
        )
    )
    note_rec2 = (
        ""
        if abs(sum_after_lines - footer_after) < 0.01
        else (
            "reconcile_required: sum_of_line_items_kWh ("
            + str(sum_after_lines)
            + ") != stated_footer ("
            + str(footer_after)
            + ")"
        )
    )
    rows.append(
        cells(
            record_type="aggregate",
            row_id="agg_after_kwh_reconcile_note",
            ratio_name="line_minus_footer_kwh_delta",
            ratio_value=sum_after_lines - footer_after,
            currency="CNY",
            notes=note_rec2,
        )
    )

    lam = 0.72
    savings = [
        (
            "save_01",
            50,
            654,
            8760,
            saved_kwh_by_row[0],
            round(saved_kwh_by_row[0] * lam, 6),
            425,
            235450,
        ),
        (
            "save_02",
            50,
            182,
            8760,
            saved_kwh_by_row[1],
            round(saved_kwh_by_row[1] * lam, 6),
            425,
            77350,
        ),
        (
            "save_03",
            80,
            282,
            8760,
            saved_kwh_by_row[2],
            round(saved_kwh_by_row[2] * lam, 6),
            580,
            163560,
        ),
        (
            "save_04",
            120,
            26,
            8760,
            saved_kwh_by_row[3],
            round(saved_kwh_by_row[3] * lam, 6),
            625,
            16250,
        ),
        (
            "save_05",
            60,
            532,
            4380,
            saved_kwh_by_row[4],
            round(saved_kwh_by_row[4] * lam, 6),
            710,
            448720,
        ),
    ]
    for sid, dpw, q, h, kwhs, fee, uc, lc in savings:
        rows.append(
            cells(
                record_type="savings_line",
                row_id=sid,
                qty_units=q,
                hours_per_year=h,
                delta_power_w=dpw,
                annual_kwh_saved_line=kwhs,
                tariff_cny_per_kwh=lam,
                annual_fee_cny=fee,
                unit_capex_cny=uc,
                line_capex_cny=lc,
                currency="CNY",
            )
        )

    total_fee_saved = total_saved_kwh * lam
    # 表尾改造后年用电量×λ = 年节费（少付）= 改造前电费 − 分项「改造后电费」合计（非「换灯后应付账单」物理量）
    annual_savings_fee_footer_basis = round(float(footer_after) * lam, 6)
    rows.append(
        cells(
            record_type="aggregate",
            row_id="agg_savings_kwh_total",
            annual_kwh_saved_line=total_saved_kwh,
            ratio_name="stated_total_annual_kwh_saved_lines_sum",
            ratio_value=total_saved_kwh,
            currency="CNY",
        )
    )
    rows.append(
        cells(
            record_type="aggregate",
            row_id="agg_line_sum_survey_retrofit_after_electricity_fee_cny",
            annual_fee_cny=round(total_fee_saved, 6),
            ratio_name="sum_savings_line_annual_fee_cny_survey_column_retrofit_after",
            ratio_value=round(total_fee_saved, 6),
            currency="CNY",
            notes=(
                "收资表「改造后电费」列：分项节电量×λ 之和 (~483636)；"
                "改造前年电费−本项=footer_after_kWh*λ (~412618)=节费/现金流基数。"
            ),
        )
    )
    rows.append(
        cells(
            record_type="aggregate",
            row_id="agg_savings_fee_total",
            annual_fee_cny=annual_savings_fee_footer_basis,
            ratio_name="stated_annual_savings_fee_cny_footer_after_kWh_times_tariff",
            ratio_value=annual_savings_fee_footer_basis,
            currency="CNY",
            notes=(
                "historical row_id agg_savings_fee_total: value = footer_after_kWh * λ = **annual savings (节费)** CNY/yr, "
                "NOT the survey「改造后电费」column (that is agg_line_sum_survey_retrofit_after_electricity_fee_cny). "
                "identity: footer_before*λ − line_sum_survey_retrofit_after = this row."
            ),
        )
    )
    rows.append(
        cells(
            record_type="aggregate",
            row_id="agg_capex_total",
            line_capex_cny=941330,
            ratio_name="stated_total_retrofit_capex_cny",
            ratio_value=941330,
            currency="CNY",
        )
    )

    before_fee_footer = float(footer_before) * lam
    ratio_delta_kwh = (footer_before - footer_after) / footer_before
    # 表尾剩余电量对应电费 / 改造前电费 ≈ 0.46（剩余负荷占比）；节费占比 = 1 − 本值（≈0.54）
    ratio_remaining_footer_fee_vs_before = annual_savings_fee_footer_basis / before_fee_footer

    rows.append(
        cells(
            record_type="aggregate",
            row_id="agg_ratio_kwh_saved_footer_totals",
            ratio_name="delta_kWh_using_stated_before_after_totals",
            ratio_value=ratio_delta_kwh,
            currency="CNY",
            notes=f"({footer_before}-{footer_after})/{footer_before}",
        )
    )
    rows.append(
        cells(
            record_type="aggregate",
            row_id="agg_ratio_fee_saved_vs_theoretical_before_fee",
            ratio_name="footer_after_fee_over_footer_before_fee",
            ratio_value=ratio_remaining_footer_fee_vs_before,
            currency="CNY",
            notes=(
                "(footer_after_kWh*λ)/(footer_before_kWh*λ)=remaining share of baseline bill after retrofit (~0.46); "
                "NOT sum(savings_line fees)/before."
            ),
        )
    )
    rows.append(
        cells(
            record_type="aggregate",
            row_id="agg_ratio_energy_saved_lines_vs_before_footer",
            ratio_name="annual_kWh_saved_sum_lines_div_before_kWh_footer",
            ratio_value=total_saved_kwh / footer_before,
            currency="CNY",
        )
    )

    rows.append(
        cells(
            record_type="om_assumption",
            row_id="om_benchmark_rate",
            ratio_name="baseline_yuan_per_lamp_year_municipal_benchmark",
            ratio_value=120,
            currency="CNY",
            notes="市政项目惯例口径，非本项目发票；用于测算基准",
        )
    )
    rows.append(
        cells(
            record_type="om_assumption",
            row_id="om_lamp_count",
            ratio_name="total_lamps_assumed_same_as_retrofit_count",
            ratio_value=1676,
            currency="CNY",
        )
    )
    rows.append(
        cells(
            record_type="om_assumption",
            row_id="om_annual_total",
            ratio_name="annual_om_cost_cny_total",
            ratio_value=201120,
            currency="CNY",
            notes="1676 * 120 = 201120; total baseline annual O&M",
        )
    )

    for comp, amt in [("人工巡检", 40), ("维修", 50), ("备件", 30)]:
        rows.append(
            cells(
                record_type="om_assumption",
                row_id=f"om_split_{comp}",
                currency="CNY",
                om_component_zh=comp,
                om_cny_per_lamp_year=amt,
                notes="示例分摊：人工巡检40 + 维修50 + 备件30 = 120 元/盏·年（占位分解，非发票明细）",
            )
        )

    for r in rows:
        assert len(r) == len(HDR), (len(r), r[:5])

    with OUT.open("w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(HDR)
        w.writerows(rows)

    print("Wrote", OUT)
    print("before lines sum", sum_before_lines, "footer", footer_before)
    print("after lines sum", sum_after_lines, "footer", footer_after)


if __name__ == "__main__":
    main()
