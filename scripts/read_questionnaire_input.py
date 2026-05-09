#!/usr/bin/env python3
"""Flatten questionnaire_01_input.xlsx to JSON for agents.

Expects sheets **填写** and/or **English** with header row 1 and columns:
  A=#  B=question text  C=unit  D=answer

Default input path (when run from monorepo root): questionnaire/questionnaire_01_input.xlsx
Portable: pass --input explicitly when this script is copied beside a skill repo without the full tree.

Usage::

    python3 scripts/read_questionnaire_input.py
    python3 scripts/read_questionnaire_input.py --input path/to/questionnaire_01_input.xlsx --sheet 填写
    python3 scripts/read_questionnaire_input.py --answers-only --compact
"""

from __future__ import annotations

import argparse
import json
import sys
from datetime import date, datetime
from pathlib import Path
from typing import Any

from openpyxl import load_workbook


def _json_safe(value: Any) -> Any:
    if value is None:
        return None
    if isinstance(value, (datetime, date)):
        return value.isoformat()
    if isinstance(value, float):
        if value.is_integer():
            return int(value)
        return value
    if isinstance(value, int):
        return value
    if isinstance(value, str):
        s = value.strip()
        return s if s else None
    return str(value)


def _read_sheet(ws) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    for r in range(2, ws.max_row + 1):
        qid = ws.cell(r, 1).value
        if qid is None or str(qid).strip() == "":
            continue
        rows.append(
            {
                "id": str(qid).strip(),
                "question": _json_safe(ws.cell(r, 2).value),
                "unit": _json_safe(ws.cell(r, 3).value),
                "answer": _json_safe(ws.cell(r, 4).value),
            }
        )
    return rows


def flatten_workbook(path: Path, sheets: list[str]) -> dict[str, Any]:
    wb = load_workbook(path, data_only=True)
    out: dict[str, Any] = {"source": str(path.resolve()), "sheets": {}}
    for name in sheets:
        if name not in wb.sheetnames:
            raise ValueError(f"Sheet {name!r} not in workbook; have: {wb.sheetnames}")
        out["sheets"][name] = {"rows": _read_sheet(wb[name])}
    return out


def answers_dict(flat: dict[str, Any], sheet_name: str) -> dict[str, Any]:
    d: dict[str, Any] = {}
    for row in flat["sheets"][sheet_name]["rows"]:
        qid = row["id"]
        if qid in d:
            sys.stderr.write(f"warning: duplicate id {qid!r} in sheet {sheet_name!r}; using last\n")
        d[qid] = row["answer"]
    return d


def main(argv: list[str] | None = None) -> int:
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument(
        "--input",
        "-i",
        type=Path,
        default=None,
        help="Path to questionnaire_01_input.xlsx (default: questionnaire/questionnaire_01_input.xlsx under cwd)",
    )
    p.add_argument(
        "--sheet",
        "-s",
        choices=("填写", "English", "both"),
        default="both",
        help="Which data sheet(s) to export (default: both)",
    )
    p.add_argument(
        "--answers-only",
        action="store_true",
        help="Emit only { id: answer } from the first exported sheet (填写 if present in workbook)",
    )
    p.add_argument(
        "--compact",
        action="store_true",
        help="Minified JSON (no indent)",
    )
    args = p.parse_args(argv)

    path = args.input
    if path is None:
        path = Path("questionnaire/questionnaire_01_input.xlsx")
    path = path.expanduser().resolve()
    if not path.is_file():
        print(f"error: file not found: {path}", file=sys.stderr)
        return 1

    if args.sheet == "both":
        sheets = ["填写", "English"]
    else:
        sheets = [args.sheet]

    try:
        flat = flatten_workbook(path, sheets)
    except ValueError as e:
        print(f"error: {e}", file=sys.stderr)
        return 1

    if args.answers_only:
        primary = "填写" if "填写" in flat["sheets"] else sheets[0]
        payload = answers_dict(flat, primary)
    else:
        payload = flat

    indent = None if args.compact else 2
    json.dump(payload, sys.stdout, ensure_ascii=False, indent=indent)
    sys.stdout.write("\n")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
