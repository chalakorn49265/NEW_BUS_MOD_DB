"""
Lightweight Streamlit entrypoint for the Excel-workbook viewer (CN).

Why this exists:
- The repo also contains a separate, larger Streamlit app in `streamlit_app.py`
  (EMC institutional model cockpit).
- For deployment/perf, we keep the workbook viewer runnable as a standalone app
  with a minimal import surface.
"""

from __future__ import annotations

import sys
from pathlib import Path

_ROOT = Path(__file__).resolve().parent
if str(_ROOT) not in sys.path:
    sys.path.insert(0, str(_ROOT))


def main() -> None:
    # Import inside main to keep cold-start light.
    # The viewer page lives under `pages/04_Tier_Comparison_Dashboard.py` (not a valid module name),
    # so we load it by file path.
    import importlib.util

    page_path = _ROOT / "pages" / "04_Tier_Comparison_Dashboard.py"
    spec = importlib.util.spec_from_file_location("tier_viewer_page", page_path)
    if spec is None or spec.loader is None:
        raise RuntimeError(f"Failed to load viewer page: {page_path}")
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    if not hasattr(mod, "main"):
        raise RuntimeError("Viewer page missing `main()`")
    mod.main()


if __name__ == "__main__":
    main()

