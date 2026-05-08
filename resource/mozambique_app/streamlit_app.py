"""
Standalone Streamlit app: Mozambique AI+Solar LaaS dashboard.

This is intentionally isolated so Streamlit multipage discovery only includes `mozambique_app/pages/*`.
"""

from __future__ import annotations

import sys
from pathlib import Path

import streamlit as st

_ROOT = Path(__file__).resolve().parents[1]  # repo root
if str(_ROOT) not in sys.path:
    sys.path.insert(0, str(_ROOT))

st.set_page_config(page_title="Mozambique — AI+Solar LaaS", layout="wide", initial_sidebar_state="expanded")

st.title("Mozambique — AI+Solar Lighting (LaaS)")
st.caption("Open the left sidebar to navigate pages: Pitch, Deal Splits, Audit Model, Export.")

