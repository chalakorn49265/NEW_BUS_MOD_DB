from __future__ import annotations

import csv
import re
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable

from business_model_comparison.provenance import Provenance, SeriesWithProv, SourceRef


_YEAR_RE = re.compile(r"第(\d+)年")


def _to_float(cell: str) -> float | None:
    s = (cell or "").strip()
    if s == "" or s == "-":
        return None
    s = s.replace(",", "")
    s = s.replace("，", "")
    s = s.replace(" ", "")
    s = s.replace("\u00a0", "")
    s = s.replace("％", "%")
    if s.endswith("%"):
        try:
            return float(s[:-1]) / 100.0
        except Exception:
            return None
    try:
        return float(s)
    except Exception:
        return None


def _read_csv_rows(path: str | Path) -> list[list[str]]:
    p = Path(path)
    with p.open("r", encoding="utf-8-sig", newline="") as f:
        reader = csv.reader(f)
        return [[c for c in row] for row in reader]


def _find_year_header_row(rows: list[list[str]]) -> tuple[int, dict[int, int]]:
    """Return (row_index, year->col_index)."""
    best: tuple[int, dict[int, int]] | None = None
    for i, row in enumerate(rows):
        mapping: dict[int, int] = {}
        for j, cell in enumerate(row):
            m = _YEAR_RE.search(cell or "")
            if m:
                mapping[int(m.group(1))] = int(j)
        if mapping:
            if best is None or len(mapping) > len(best[1]):
                best = (i, mapping)
    if best is None:
        raise ValueError("Could not find a header row containing year labels like 第1年.")
    return best


def _extract_row_year_values(
    rows: list[list[str]],
    *,
    row_matcher: callable,
    year_to_col: dict[int, int],
    unit: str,
    source_file: str,
    row_label: str,
    transform: str,
) -> SeriesWithProv:
    for row in rows:
        if row_matcher(row):
            out: dict[int, float] = {}
            for y, col in year_to_col.items():
                if col < len(row):
                    v = _to_float(row[col])
                    if v is not None:
                        out[int(y)] = float(v)
            return SeriesWithProv(
                values_by_year=out,
                provenance=Provenance(
                    sources=(SourceRef(file=source_file, row_label=row_label),),
                    units=unit,
                    transform=transform,
                ),
            )
    raise ValueError(f"Could not find row for label: {row_label}")


@dataclass(frozen=True)
class RoadlightParsed:
    """Normalized annual inputs (all monetary values in RMB unless otherwise noted)."""

    years: list[int]
    baseline_revenue_trust_fee_rmb_y: SeriesWithProv  # what client pays; also provider revenue in baseline
    baseline_opex_cash_rmb_y: SeriesWithProv  # cash opex (ex depreciation)
    baseline_opex_electricity_rmb_y: SeriesWithProv  # electricity component (subset of cash opex)
    depreciation_rmb_y: SeriesWithProv  # accounting depreciation
    debt_interest_rmb_y: SeriesWithProv
    debt_principal_rmb_y: SeriesWithProv
    capex_cash_rmb_y0: float
    capex_provenance: Provenance


def load_income_analysis(path: str | Path) -> SeriesWithProv:
    rows = _read_csv_rows(path)
    _hdr_idx, year_to_col = _find_year_header_row(rows)

    def match(row: list[str]) -> bool:
        return any((c or "").strip() == "托管收入" for c in row)

    return _extract_row_year_values(
        rows,
        row_matcher=match,
        year_to_col=year_to_col,
        unit="10k_RMB_per_year",
        source_file=str(path),
        row_label="托管收入",
        transform="Parsed annual 托管收入 values from year columns 第1年..; unit 万元/year as provided.",
    )


def load_opex_cash_and_depr(path: str | Path) -> tuple[SeriesWithProv, SeriesWithProv, SeriesWithProv]:
    rows = _read_csv_rows(path)
    _hdr_idx, year_to_col = _find_year_header_row(rows)

    def row_label_equals(label: str) -> callable:
        def _m(row: list[str]) -> bool:
            return any((c or "").strip() == label for c in row)

        return _m

    electricity = _extract_row_year_values(
        rows,
        row_matcher=row_label_equals("改造后电费"),
        year_to_col=year_to_col,
        unit="10k_RMB_per_year",
        source_file=str(path),
        row_label="改造后电费",
        transform="Parsed annual 改造后电费 from 第1年.. columns; unit 万元/year.",
    )
    staff = _extract_row_year_values(
        rows,
        row_matcher=row_label_equals("职工薪酬费用"),
        year_to_col=year_to_col,
        unit="10k_RMB_per_year",
        source_file=str(path),
        row_label="职工薪酬费用",
        transform="Parsed annual 职工薪酬费用 from 第1年.. columns; unit 万元/year.",
    )
    materials = _extract_row_year_values(
        rows,
        row_matcher=row_label_equals("维修材料费"),
        year_to_col=year_to_col,
        unit="10k_RMB_per_year",
        source_file=str(path),
        row_label="维修材料费",
        transform="Parsed annual 维修材料费 from 第1年.. columns; unit 万元/year.",
    )
    vehicles = _extract_row_year_values(
        rows,
        row_matcher=row_label_equals("车辆费用"),
        year_to_col=year_to_col,
        unit="10k_RMB_per_year",
        source_file=str(path),
        row_label="车辆费用",
        transform="Parsed annual 车辆费用 from 第1年.. columns; unit 万元/year.",
    )
    mgmt = _extract_row_year_values(
        rows,
        row_matcher=row_label_equals("管理费用"),
        year_to_col=year_to_col,
        unit="10k_RMB_per_year",
        source_file=str(path),
        row_label="管理费用",
        transform="Parsed annual 管理费用 from 第1年.. columns; unit 万元/year.",
    )
    depr = _extract_row_year_values(
        rows,
        row_matcher=row_label_equals("设备折旧费"),
        year_to_col=year_to_col,
        unit="10k_RMB_per_year",
        source_file=str(path),
        row_label="设备折旧费",
        transform="Parsed annual 设备折旧费 (accounting depreciation) from 第1年.. columns; unit 万元/year.",
    )

    cash = SeriesWithProv(
        values_by_year={y: electricity.get(y) + staff.get(y) + materials.get(y) + vehicles.get(y) + mgmt.get(y) for y in sorted(year_to_col.keys())},
        provenance=Provenance(
            sources=(
                SourceRef(file=str(path), row_label="改造后电费"),
                SourceRef(file=str(path), row_label="职工薪酬费用"),
                SourceRef(file=str(path), row_label="维修材料费"),
                SourceRef(file=str(path), row_label="车辆费用"),
                SourceRef(file=str(path), row_label="管理费用"),
            ),
            units="10k_RMB_per_year",
            transform="cash_OPEX = electricity + staff + materials + vehicles + mgmt; excludes depreciation.",
        ),
    )
    return cash, electricity, depr


def load_loan_schedule(path: str | Path) -> tuple[SeriesWithProv, SeriesWithProv]:
    rows = _read_csv_rows(path)
    # Loan file is tidy-ish: year rows begin with an integer.
    principal: dict[int, float] = {}
    interest: dict[int, float] = {}
    for row in rows:
        if not row:
            continue
        y = _to_float(row[0])
        if y is None:
            continue
        y_i = int(y)
        if y_i <= 0:
            continue
        # Column meanings per header:
        # 0 年度期数, 1 计息贷款本金(万元), 2 期末还款计划(万元), 3 利率, 4 年度利息(万元)
        pmt = _to_float(row[2]) if len(row) > 2 else None
        intr = _to_float(row[4]) if len(row) > 4 else None
        if pmt is not None:
            principal[y_i] = float(pmt)
        if intr is not None:
            interest[y_i] = float(intr)

    p_series = SeriesWithProv(
        values_by_year=principal,
        provenance=Provenance(
            sources=(SourceRef(file=str(path), row_label="期末还款计划（万元）"),),
            units="10k_RMB_per_year",
            transform="Parsed debt principal repayments by year from loan schedule; unit 万元/year.",
        ),
    )
    i_series = SeriesWithProv(
        values_by_year=interest,
        provenance=Provenance(
            sources=(SourceRef(file=str(path), row_label="年度利息（万元）"),),
            units="10k_RMB_per_year",
            transform="Parsed debt interest by year from loan schedule; unit 万元/year.",
        ),
    )
    return p_series, i_series


def load_capex_total_investment(path: str | Path) -> tuple[float, Provenance]:
    rows = _read_csv_rows(path)
    # Find row containing "总投资" and take last numeric cell.
    for row in rows:
        if any((c or "").strip() == "总投资" for c in row):
            nums: list[float] = []
            for cell in row:
                v = _to_float(cell)
                if v is not None:
                    nums.append(float(v))
            if not nums:
                break
            total_10k = float(nums[-1])  # 万元
            prov = Provenance(
                sources=(SourceRef(file=str(path), row_label="总投资", notes="Used as cash CAPEX source of truth (万元)."),),
                units="10k_RMB",
                transform="CAPEX_cash_y0 = capex.csv:总投资 (万元).",
            )
            return total_10k * 10_000.0, prov
    raise ValueError("Could not find CAPEX total investment (row label 总投资).")


def load_depreciation_schedule(path: str | Path) -> SeriesWithProv:
    rows = _read_csv_rows(path)
    _hdr_idx, year_to_col = _find_year_header_row(rows)

    def match(row: list[str]) -> bool:
        return any((c or "").strip() == "年折旧额" for c in row)

    return _extract_row_year_values(
        rows,
        row_matcher=match,
        year_to_col=year_to_col,
        unit="10k_RMB_per_year",
        source_file=str(path),
        row_label="年折旧额",
        transform="Parsed annual depreciation expense from 年折旧额 row; unit 万元/year.",
    )


def load_roadlight_all(data_dir: str | Path) -> RoadlightParsed:
    d = Path(data_dir)
    income = load_income_analysis(d / "income_analysis.csv")
    opex_cash, opex_elec, depr_opex = load_opex_cash_and_depr(d / "opex.csv")
    debt_principal, debt_interest = load_loan_schedule(d / "loan.csv")
    capex_y0_rmb, capex_prov = load_capex_total_investment(d / "capex.csv")
    depr_sched = load_depreciation_schedule(d / "product_depreciation overtime.csv")

    years = sorted(set(income.years()) | set(opex_cash.years()) | set(depr_opex.years()) | set(depr_sched.years()))
    years = [y for y in years if 1 <= y <= 15]

    # Use depreciation from opex.csv for accounting P&L (aligns with baseline sheet).
    # Keep schedule as a cross-check; surfaced in dashboard traceability.
    depr = SeriesWithProv(
        values_by_year=depr_opex.reindex_years(years).values_by_year,
        provenance=Provenance(
            sources=(
                *depr_opex.provenance.sources,
                SourceRef(file=str(d / "product_depreciation overtime.csv"), row_label="年折旧额", notes="Cross-check schedule."),
            ),
            units="10k_RMB_per_year",
            transform="Use opex.csv depreciation as accounting depreciation; cross-check against depreciation schedule.",
        ),
    )

    return RoadlightParsed(
        years=years,
        baseline_revenue_trust_fee_rmb_y=income.reindex_years(years),
        baseline_opex_cash_rmb_y=opex_cash.reindex_years(years),
        baseline_opex_electricity_rmb_y=opex_elec.reindex_years(years),
        depreciation_rmb_y=depr.reindex_years(years),
        debt_interest_rmb_y=debt_interest.reindex_years(years),
        debt_principal_rmb_y=debt_principal.reindex_years(years),
        capex_cash_rmb_y0=float(capex_y0_rmb),
        capex_provenance=capex_prov,
    )


def to_rmb(series_10k: SeriesWithProv) -> SeriesWithProv:
    return SeriesWithProv(
        values_by_year={y: v * 10_000.0 for y, v in series_10k.values_by_year.items()},
        provenance=Provenance(
            sources=series_10k.provenance.sources,
            units="RMB_per_year",
            transform=series_10k.provenance.transform + " Then multiplied by 10,000 to convert 万元→元(RMB).",
        ),
    )


def common_years(*series: Iterable[SeriesWithProv]) -> list[int]:
    ys: set[int] | None = None
    for s in series:
        for ss in s:
            if ys is None:
                ys = set(ss.values_by_year.keys())
            else:
                ys &= set(ss.values_by_year.keys())
    return sorted(ys or [])

