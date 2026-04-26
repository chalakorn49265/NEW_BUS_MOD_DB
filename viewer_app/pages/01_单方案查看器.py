from __future__ import annotations

import importlib.util
import sys
from pathlib import Path

_REPO_ROOT = Path(__file__).resolve().parents[2]
if str(_REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(_REPO_ROOT))


def _load_viewer_module():
    page_path = _REPO_ROOT / "pages" / "04_Tier_Comparison_Dashboard.py"
    spec = importlib.util.spec_from_file_location("tier_viewer_page", page_path)
    if spec is None or spec.loader is None:
        raise RuntimeError(f"Failed to load viewer page: {page_path}")
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


def main() -> None:
    mod = _load_viewer_module()
    if not hasattr(mod, "main"):
        raise RuntimeError("Viewer page missing `main()`")
    mod.main()


if __name__ == "__main__":
    main()

