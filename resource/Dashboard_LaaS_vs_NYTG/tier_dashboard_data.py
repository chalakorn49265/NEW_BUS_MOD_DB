from __future__ import annotations

from dataclasses import asdict
from pathlib import Path

import pandas as pd

from Dashboard_LaaS_vs_NYTG.workbook_extract import TierExtract, discover_and_extract


def _safe_sum(xs: list[float | None]) -> float | None:
    if any(x is None for x in xs):
        return None
    return float(sum(float(x) for x in xs if x is not None))


def _safe_min(xs: list[float | None]) -> float | None:
    vals = [float(x) for x in xs if x is not None]
    return float(min(vals)) if vals else None


def _safe_y1(xs: list[float | None]) -> float | None:
    return float(xs[0]) if xs and xs[0] is not None else None


def build_tier_tables(*, new_models_dir: Path) -> tuple[pd.DataFrame, pd.DataFrame, list[TierExtract]]:
    tiers = discover_and_extract(new_models_dir=new_models_dir)

    # Summary: one row per tier
    rows = []
    for t in tiers:
        rows.append(
            {
                "tier": t.tier_name,
                "file_path": t.file_path,
                "term_years": t.term_years,
                "upfront": t.upfront,
                "tail_discount": t.tail_discount,
                "laas_fee_y1_input": t.laas_fee_y1_input,
                "emc_fee_y1": t.emc_fee_y1,
                "emc_owner_pays_elec_flag": t.emc_owner_pays_elec_flag,
                "lamps": t.lamps,
                "capex_emc_per_lamp": t.capex_emc_per_lamp,
                "capex_laas_per_lamp": t.capex_laas_per_lamp,
                "baseline_electricity_y1": t.baseline_electricity_y1,
                "electricity_price_per_kwh": t.electricity_price_per_kwh,
                "watts_per_lamp": t.watts_per_lamp,
                "hours_per_day": t.hours_per_day,
                "days_per_year": t.days_per_year,
                "emc_saving_rate": t.emc_saving_rate,
                "laas_saving_rate": t.laas_saving_rate,
                "opex_om_per_lamp": t.opex_om_per_lamp,
                "opex_platform": t.opex_platform,
                "opex_spares": t.opex_spares,
                "laas_opex_om_per_lamp": t.laas_opex_om_per_lamp,
                "laas_opex_platform": t.laas_opex_platform,
                "laas_opex_spares": t.laas_opex_spares,
                "product_key": t.product_key,
                # Provider KPIs
                "emc_npv": t.emc_npv,
                "emc_irr": t.emc_irr,
                "laas_npv": t.laas_npv,
                "laas_irr": t.laas_irr,
                "delta_npv": (t.laas_npv - t.emc_npv) if (t.laas_npv is not None and t.emc_npv is not None) else None,
                "delta_irr": (t.laas_irr - t.emc_irr) if (t.laas_irr is not None and t.emc_irr is not None) else None,
                # Owner metrics
                "owner_save_emc_y1": _safe_y1(t.owner_save_emc_y),
                "owner_save_laas_y1": _safe_y1(t.owner_save_laas_y),
                "owner_save_delta_y1": (
                    _safe_y1(t.owner_save_laas_y) - _safe_y1(t.owner_save_emc_y)
                    if (_safe_y1(t.owner_save_laas_y) is not None and _safe_y1(t.owner_save_emc_y) is not None)
                    else None
                ),
                "owner_save_emc_min": _safe_min(t.owner_save_emc_y),
                "owner_save_laas_min": _safe_min(t.owner_save_laas_y),
                "owner_save_emc_sum10": _safe_sum(t.owner_save_emc_y),
                "owner_save_laas_sum10": _safe_sum(t.owner_save_laas_y),
                "owner_spend_emc_sum10": _safe_sum(t.owner_spend_emc_y),
                "owner_spend_laas_sum10": _safe_sum(t.owner_spend_laas_y),
                "provider_lines": t.provider_lines,
            }
        )

    df_summary = pd.DataFrame(rows).sort_values("tier").reset_index(drop=True)

    # Long form: year series
    series_rows = []
    for t in tiers:
        for i, y in enumerate(range(1, 11)):
            series_rows.append(
                {
                    "tier": t.tier_name,
                    "year": y,
                    "owner_spend_emc": t.owner_spend_emc_y[i] if i < len(t.owner_spend_emc_y) else None,
                    "owner_spend_laas": t.owner_spend_laas_y[i] if i < len(t.owner_spend_laas_y) else None,
                    "owner_save_emc": t.owner_save_emc_y[i] if i < len(t.owner_save_emc_y) else None,
                    "owner_save_laas": t.owner_save_laas_y[i] if i < len(t.owner_save_laas_y) else None,
                    "laas_fee": t.laas_fee_y[i] if i < len(t.laas_fee_y) else None,
                }
            )

    df_long = pd.DataFrame(series_rows).sort_values(["tier", "year"]).reset_index(drop=True)
    return df_summary, df_long, tiers


def tier_traceability_dict(t: TierExtract) -> dict:
    d = asdict(t)
    d["cell_sources"] = {
        "provider_kpis": {
            "emc_npv": "01_Dashboard!C19",
            "emc_irr": "01_Dashboard!C20",
            "laas_npv": "01_Dashboard!D19",
            "laas_irr": "01_Dashboard!D20",
        },
        "owner_series": {
            "emc_spend": "05_Annual_Model!D34:M34",
            "emc_net_savings": "05_Annual_Model!D35:M35",
            "laas_spend": "05_Annual_Model!D61:M61",
            "laas_net_savings": "05_Annual_Model!D62:M62",
        },
        "laas_fee_schedule": "05_Annual_Model!D40:M40",
        "inputs": {
            "term_years": "02_Inputs!D5",
            "laas_fee_y1": "02_Inputs!D29",
            "upfront": "02_Inputs!D45",
            "tail_discount": "02_Inputs!D46",
            "emc_fee_y1": "02_Inputs!D19",
            "emc_owner_pays_elec_flag": "02_Inputs!D25",
            "lamps": "02_Inputs!D6",
            "capex_emc_per_lamp": "02_Inputs!D18",
            "capex_laas_per_lamp": "02_Inputs!D28",
            "baseline_electricity_y1": "03_Baseline!D12",
            "emc_saving_rate": "02_Inputs!D21",
            "laas_saving_rate": "02_Inputs!D31",
            "opex_om_per_lamp": "02_Inputs!D22",
            "opex_platform": "02_Inputs!D23",
            "opex_spares": "02_Inputs!D24",
            "laas_opex_om_per_lamp": "02_Inputs!D32",
            "laas_opex_platform": "02_Inputs!D33",
            "laas_opex_spares": "02_Inputs!D34",
            "product_key": "02_Inputs!D48",
        },
        "provider_cashflow_block": {
            "labels": "05_Annual_Model (通过文本标签定位行；读取 C..M = Y0..Y10 缓存值)",
        },
    }
    return d

