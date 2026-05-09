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
                "source=user spreadsheet screenshot (retrofit table); theoretical_fee_tariff=0.72 CNY/kWh; "
                "street_lamps=532 here (differs from earlier survey XLS 632); reconciliations flagged where "
                "line sums != footer totals."
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

    after_lines = [
        ("after_01", "LED隧道灯", "LED tunnel", 40, 654, 8760, 194122),
        ("after_02", "LED隧道灯", "LED tunnel", 50, 182, 8760, 132596),
        ("after_03", "LED隧道灯", "LED tunnel", 100, 282, 8760, 247032),
        ("after_04", "LED隧道灯", "LED tunnel", 200, 26, 8760, 45552),
        ("after_05", "LED路灯", "LED street", 120, 532, 4380, 332179),
    ]
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
    footer_after = 948430

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

    savings = [
        ("save_01", 50, 654, 8760, 145591.2, 104825.664, 425, 235450),
        ("save_02", 50, 182, 8760, 79716.0, 57395.52, 425, 77350),
        ("save_03", 80, 282, 8760, 197625.6, 142290.432, 580, 163560),
        ("save_04", 120, 26, 8760, 27331.2, 19678.464, 625, 16250),
        ("save_05", 60, 532, 4380, 221452.8, 159446.016, 710, 448720),
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
                tariff_cny_per_kwh=0.72,
                annual_fee_cny=fee,
                unit_capex_cny=uc,
                line_capex_cny=lc,
                currency="CNY",
            )
        )

    rows.append(
        cells(
            record_type="aggregate",
            row_id="agg_savings_kwh_total",
            annual_kwh_saved_line=671716.8,
            ratio_name="stated_total_annual_kwh_saved_lines_sum",
            ratio_value=671716.8,
            currency="CNY",
        )
    )
    rows.append(
        cells(
            record_type="aggregate",
            row_id="agg_savings_fee_total",
            annual_fee_cny=483636.096,
            ratio_name="stated_total_annual_fee_saved_cny",
            ratio_value=483636.096,
            currency="CNY",
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

    before_fee_footer = footer_before * 0.72
    ratio_delta_kwh = (footer_before - footer_after) / footer_before
    ratio_fee_saved_vs_before_fee = (before_fee_footer - 483636.096) / before_fee_footer

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
            ratio_name="fee_savings_vs_before_kWh_times_tariff",
            ratio_value=ratio_fee_saved_vs_before_fee,
            currency="CNY",
            notes="matches ~46pct narrative when baseline fee = footer_before_kWh * 0.72",
        )
    )
    rows.append(
        cells(
            record_type="aggregate",
            row_id="agg_ratio_energy_saved_lines_vs_before_footer",
            ratio_name="annual_kWh_saved_sum_lines_div_before_kWh_footer",
            ratio_value=671716.8 / footer_before,
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
