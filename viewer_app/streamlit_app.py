"""
Standalone Streamlit app: Workbook Viewer (CN).

This folder is intentionally isolated so Streamlit multipage discovery only
includes `viewer_app/pages/*` (and does NOT pull in the repo-root `pages/*`).
"""

from __future__ import annotations

import sys
from pathlib import Path

import streamlit as st

_ROOT = Path(__file__).resolve().parents[1]  # repo root
if str(_ROOT) not in sys.path:
    sys.path.insert(0, str(_ROOT))

st.set_page_config(page_title="EMC → LaaS 工作簿查看器（CN）", layout="wide", initial_sidebar_state="expanded")

st.title("能源托管/EMC → LaaS：工作簿查看器（独立版）")
st.caption("请从左侧选择页面：`01_单方案查看器`。该应用仅用于展示生成的 Excel 工作簿（new_models）。")

