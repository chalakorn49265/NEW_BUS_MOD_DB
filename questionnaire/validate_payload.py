from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import jsonschema
from jsonschema import Draft202012Validator

_ROOT = Path(__file__).resolve().parents[1]
SCHEMA_PATH_V2026_01 = _ROOT / "schemas" / "project_capex_pack.v2026_01.schema.json"


def _load_schema() -> dict[str, Any]:
    with open(SCHEMA_PATH_V2026_01, encoding="utf-8") as f:
        return json.load(f)


def validate_project_capex_pack(payload: dict[str, Any], *, schema_path: Path | None = None) -> None:
    """
    Validate payload against the packaged JSON Schema. Raises jsonschema.ValidationError on failure.
    """
    sp = schema_path or SCHEMA_PATH_V2026_01
    with open(sp, encoding="utf-8") as f:
        schema = json.load(f)
    Draft202012Validator.check_schema(schema)
    validator = Draft202012Validator(schema)
    validator.validate(payload)


def validate_file(path: Path) -> None:
    with open(path, encoding="utf-8") as f:
        payload = json.load(f)
    validate_project_capex_pack(payload)
