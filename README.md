# EMC institutional financial model (Python)

Standalone decision engine aligned with the MOZ v2 Excel workbook logic (see `DELIVERABLES/TRACK_C/build_moz_institutional_model.py` for formula parity). **This package does not read Excel files.**

## Setup

```bash
cd EMC_INSTITUTIONAL_MODEL
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

## Tests

```bash
pytest tests/ -q
```

## Streamlit

```bash
streamlit run streamlit_app.py
```

Use a dedicated virtual environment; do not reuse `DELIVERABLES/TRACK_C/.venv`.