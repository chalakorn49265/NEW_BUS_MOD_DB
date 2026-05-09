#!/usr/bin/env bash
# Single-page Wenzhou dashboard only (does not load repo-root pages/).
cd "$(dirname "$0")"
exec python3 -m streamlit run wenzhou/run_dashboard.py "$@"
