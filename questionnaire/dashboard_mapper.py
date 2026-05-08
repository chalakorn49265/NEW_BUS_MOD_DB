from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any, Mapping

# Scenario order for final narrative (LaaS last for emphasis)
SCENARIO_ORDER = ("baseline", "normal_led", "emc", "laas")

OM_CHART_KEYS = (
    "labor",
    "materials",
    "other",
    "inspection",
    "cleaning",
    "testing",
    "platform_software",
    "spares",
    "battery_reserve",
    "total_om",
)


@dataclass
class ChartBundle:
    """Structured inputs for Plotly/Streamlit or agent-generated briefs."""

    payback: dict[str, Any]
    opex_comparison: dict[str, Any]
    maintenance_breakdown: dict[str, Any]
    capex_summary: dict[str, Any]
    flags: dict[str, bool] = field(default_factory=dict)
    raw: dict[str, Any] = field(default_factory=dict)


def _safe_float(x: object, default: float = 0.0) -> float:
    try:
        v = float(x)  # type: ignore[arg-type]
        return v if v == v else default
    except Exception:
        return default


def build_dashboard_chart_bundle(payload: Mapping[str, Any]) -> ChartBundle:
    """
    Map a validated `project_capex_pack` to chart-friendly structures.
    Does not re-validate; call validate_project_capex_pack first.
    """
    ct = payload.get("capex_triplet") or {}
    scale = payload.get("scale") or {}
    opex = payload.get("opex_annual_by_scenario") or {}
    om = payload.get("maintenance_breakdown_by_scenario") or {}
    laas = payload.get("commercial_laas") or {}
    kpis = payload.get("calculated_kpis") or {}

    lights = int(scale.get("number_of_lights") or 0)

    capex_summary = {
        "capex_ours": _safe_float(ct.get("capex_ours")),
        "capex_baseline_incumbent": _safe_float(ct.get("capex_baseline_incumbent")),
        "capex_normal_led": _safe_float(ct.get("capex_normal_led")),
        "currency": str(ct.get("currency") or "USD"),
        "per_light_ours": (_safe_float(ct.get("capex_ours")) / lights) if lights else None,
        "per_light_baseline": (_safe_float(ct.get("capex_baseline_incumbent")) / lights) if lights else None,
        "per_light_normal_led": (_safe_float(ct.get("capex_normal_led")) / lights) if lights else None,
    }

    # Payback: prefer explicit KPI; else rough months from LaaS fee vs our CAPEX if possible
    payback_years = kpis.get("payback_years")
    payback: dict[str, Any] = {
        "payback_years_reported": payback_years,
        "payback_months_reported": (float(payback_years) * 12.0) if payback_years is not None else None,
        "assumption_note": "If payback_years missing, run full cashflow model or collect commercial_laas + opex series.",
    }
    fee = _safe_float(laas.get("annual_service_fee"))
    capex_ours = _safe_float(ct.get("capex_ours"))
    if payback_years is None and fee > 0 and capex_ours > 0:
        # Crude static payback on fee alone (illustrative; not replacement for project cashflow)
        payback["payback_years_static_fee_only"] = capex_ours / fee
        payback["flags_fee_only_approximation"] = True

    # OPEX comparison
    scenarios_present = [s for s in SCENARIO_ORDER if s in opex and isinstance(opex[s], Mapping)]
    opex_comparison = {
        "scenarios": scenarios_present,
        "series_total": {s: _safe_float((opex[s] or {}).get("total_annual")) for s in scenarios_present},
        "series_electricity": {s: _safe_float((opex[s] or {}).get("electricity_annual")) for s in scenarios_present},
        "series_non_electric": {s: _safe_float((opex[s] or {}).get("non_electric_annual")) for s in scenarios_present},
    }

    # Maintenance breakdown — stacked categories per scenario
    om_scenarios = [s for s in SCENARIO_ORDER if s in om and isinstance(om[s], Mapping)]
    breakdown_by_scenario: dict[str, dict[str, float]] = {}
    for s in om_scenarios:
        block = om[s] or {}
        breakdown_by_scenario[s] = {k: _safe_float(block.get(k)) for k in OM_CHART_KEYS if block.get(k) is not None}

    maintenance_breakdown = {
        "scenarios": om_scenarios,
        "categories": list(OM_CHART_KEYS),
        "by_scenario": breakdown_by_scenario,
    }

    flags = {
        "has_opex_by_scenario": bool(scenarios_present),
        "has_maintenance_breakdown": bool(om_scenarios),
        "needs_phase2_for_charts": not (scenarios_present and om_scenarios),
    }

    return ChartBundle(
        payback=payback,
        opex_comparison=opex_comparison,
        maintenance_breakdown=maintenance_breakdown,
        capex_summary=capex_summary,
        flags=flags,
        raw=dict(payload),
    )


def bundle_to_plotly_dicts(bundle: ChartBundle) -> dict[str, Any]:
    """
    Minimal Plotly-friendly dicts (caller supplies styling). Useful for agents/tests.
    """
    import plotly.graph_objects as go

    fig_opex = None
    if bundle.opex_comparison.get("scenarios"):
        sc = bundle.opex_comparison["scenarios"]
        y = [bundle.opex_comparison["series_total"].get(s, 0.0) for s in sc]
        fig_opex = go.Figure(data=[go.Bar(x=list(sc), y=y, name="Annual OPEX")])

    fig_om = None
    if bundle.maintenance_breakdown.get("scenarios"):
        # Single stacked bar per scenario — simplified: sum known keys except total_om double-count guard
        scenarios = bundle.maintenance_breakdown["scenarios"]
        traces = []
        # omit total_om from stack if present to avoid double count
        cats = [c for c in OM_CHART_KEYS if c != "total_om"]
        for cat in cats:
            ys = []
            for s in scenarios:
                ys.append(bundle.maintenance_breakdown["by_scenario"].get(s, {}).get(cat, 0.0))
            if sum(ys) > 0:
                traces.append(go.Bar(name=cat, x=list(scenarios), y=ys))
        if traces:
            fig_om = go.Figure(data=traces)
            fig_om.update_layout(barmode="stack")

    return {"opex_bar": fig_opex, "maintenance_stacked": fig_om}
