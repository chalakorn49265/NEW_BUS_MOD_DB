from __future__ import annotations

import json
from pathlib import Path

import pytest

from questionnaire.dashboard_mapper import build_dashboard_chart_bundle
from questionnaire.validate_payload import SCHEMA_PATH_V2026_01, validate_project_capex_pack

_ROOT = Path(__file__).resolve().parents[1]


def test_schema_examples_validate() -> None:
    ex_dir = _ROOT / "schemas" / "examples"
    for name in ("minimal_valid.v2026_01.json", "full_phase2_example.v2026_01.json"):
        with open(ex_dir / name, encoding="utf-8") as f:
            payload = json.load(f)
        validate_project_capex_pack(payload, schema_path=SCHEMA_PATH_V2026_01)


def test_mapper_builds_bundle() -> None:
    with open(_ROOT / "schemas" / "examples" / "full_phase2_example.v2026_01.json", encoding="utf-8") as f:
        payload = json.load(f)
    validate_project_capex_pack(payload)
    b = build_dashboard_chart_bundle(payload)
    assert b.flags["has_opex_by_scenario"] is True
    assert b.flags["has_maintenance_breakdown"] is True
    assert "laas" in b.opex_comparison["scenarios"]
