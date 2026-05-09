# Generalized model — data acquisition map

This guide ties **client-fixed facts** (questionnaire) to a **generalized comparison model** (incumbent vs “our” solution), and separates what stays **internal / adjustable** (sliders) from what must be **collected on site**.

**Canonical questionnaire (full spec + Definitions):** [`questionnaire/questionnaire_01.xlsx`](../questionnaire/questionnaire_01.xlsx)  
**Fast entry:** [`questionnaire/questionnaire_01_input.xlsx`](../questionnaire/questionnaire_01_input.xlsx) — sheet **填写** (中文) and sheet **English** (same row order and `#` ids); EN text from the master workbook. **Which questions appear** (and in what order) is controlled by [`questionnaire/questionnaire_input_row_ids.txt`](../questionnaire/questionnaire_input_row_ids.txt) (if missing or empty, all questions from the master are included). **Answer** column: Excel **Data Validation** — dropdown lists for categorical/binary fields; date range for **INT6**; non‑negative **whole** numbers for key counts/years; non‑negative **decimal** for money/kWh (with **C4b–C4d** allowing `N/A` or a number); long **qualitative** answers stay free text. Hidden **Lists** / **Lists_EN**. Regenerate with [`questionnaire/build_questionnaire_input_xlsx.py`](../questionnaire/build_questionnaire_input_xlsx.py).  
**Flat exports (3 列):** [`questionnaire/questionnaire_01_3col_zh.xlsx`](../questionnaire/questionnaire_01_3col_zh.xlsx) — regenerate with [`questionnaire/build_questionnaire_01_3col.py`](../questionnaire/build_questionnaire_01_3col.py)

**Optional structured dashboard pack (when you freeze numbers for JSON):** [`MODEL_INPUTS_DASHBOARD_MINIMUM.md`](MODEL_INPUTS_DASHBOARD_MINIMUM.md), schema [`schemas/project_capex_pack.v2026_01.schema.json`](../schemas/project_capex_pack.v2026_01.schema.json).

---

## Conceptual split

| Layer | Source | Examples |
|-------|--------|----------|
| **Client-fixed** | Questionnaire + bills / O&M records | Lights, kWh, tariffs, incumbent spend, payer, contract appetite |
| **Internal / adjustable** | Your product & finance assumptions | Energy savings %, labour reduction, platform fee, replacement curve, “our” CAPEX |
| **Derived / export** | Model outputs or manual consolidation | `project_capex_pack` JSON for dashboards |

---

## Acquisition flow

```mermaid
flowchart LR
  prep[Prep_INT_currency_scope]
  interview[Interview_A_through_K]
  reconcile[Reconcile_B1_vs_B2_to_B4_D2]
  model[Load_fixed_into_model]
  tune[Adjust_internal_sliders]
  prep --> interview --> reconcile --> model --> tune
```

**Reconcile** uses the **Definitions** sheet in `questionnaire_01.xlsx` (B1 vs B2–B4; D2 vs B3). **Electricity detail** uses **C4 + C4a–C4e** (flat vs TOU; peak/off-peak/shoulder **¥/kWh**; time windows or %kWh by period).

---

## Step-by-step checklist

1. **Before visit** — INT1, INT4, INT5, INT6, INT7 (and INT2, INT3 as needed): project id, currency, baseline period, date, respondent.
2. **Scope & inventory** — A1–A4, A6–A7: geography, project type, counts, fixture mix, age / failure pressure.
3. **Money baseline** — B1–B4, B6, B8–B9: totals and splits; check Definitions for double-counting.
4. **Energy** — C1–C2, **C4, C4a–C4e**, C5–C7, C9, C10–C12: bills, kWh, **explicit TOU prices when applicable**, hours, control/dimming, abnormalities, demand charges, metering, on-site generation.
5. **O&M** — D2–D10, D9: contract, unit cost, inspection, tickets, repair times, spares, delivery model.
6. **Infra / civil** — E1–E8: trenching, unit costs, per-light build, cabinets, hidden costs, underground vs overhead share.
7. **Connectivity / CMS** — F1–F2.
8. **Incumbent third-party contract (EMC / ESCO)** — **J1–J6** (placed after budget, before *new* model acceptance): whether an EMC-type contract already covers the portfolio, remaining term, payment mechanic, **who pays electricity under that contract**, material constraints, rough annual cash tied to the arrangement.
9. **Budget & future commercial structure** — G1, G11–G12 (payer, stress, evidence); H7, H10, I2, I5, K6 (ownership, term, LaaS-style acceptance, disposal, data sharing).

---

## ID → generalized model (high level)

| Question IDs | Model use |
|--------------|-----------|
| INT* | Metadata, traceability, reporting currency & baseline period |
| A3, A4 | Scale, replication upside |
| B2, B3, B6, B8 | Incumbent cost **components** (electricity, O&M, inspection, software) |
| B1 | **Sanity check** vs B2+B3+B4 |
| C2, **C4, C4a–C4e**, C5–C7 | Energy baseline & how spend maps to **price periods** |
| D* | O&M depth (tickets, times, spares, outsourced vs in-house) |
| E* | Civil / conventional CAPEX **priors** (not “our” pricing) |
| **J1–J6** | **Incumbent** EMC/ESCO coverage, term, mechanic, **electricity payer**, constraints, annual contract cash |
| G / H / I / K | Payer, risk, ownership, term, **future** structure, disposal, references |

**Internal sliders** (not on this form): savings vs baseline, fee structures, escalation, your CAPEX — document them in the model UI or internal playbook.

---

## Known questionnaire gaps (resolved for tariffs)

Previously a single **C4** bundle made **peak vs off-peak price per kWh** easy to skip. The workbook now adds:

| ID | Purpose |
|----|---------|
| **C4a** | Billing type: flat / TOU / multi-tier |
| **C4b** | Peak (or top tier) **price per kWh** |
| **C4c** | Off-peak **price per kWh** |
| **C4d** | Shoulder / mid-peak **price per kWh** (if any) |
| **C4e** | Time windows **or** approximate **% of annual kWh** per period |

Link with **C5–C7** (hours on, schedule, dimming) when mapping load to price buckets.

---

## Minimum field set (example sales strip)

If you use a **short** column list (e.g. a trimmed spreadsheet), a typical subset is: INT1, INT3, INT4, A1–A7, B1–B4, B6, B8–B9, C1–C2, C4–C7, C9, C12, D2–D8, D10, E1–E6, G1, G11, **J1, J4** (incumbent EMC? who pays power), H7, H10, I5, K6.  
The **full** workbook adds INT2, INT5–INT7, **C4a–C4e**, C10–C11, D9, E7–E8, F1–F2, **J2–J3, J5–J6**, G12, I2 — use the full set when building a defensible generalized model.

---

## Cross-links

- Feishu / Base mapping: [`feishu_field_mapping.md`](feishu_field_mapping.md)  
- Dashboard-minimum JSON fields: [`MODEL_INPUTS_DASHBOARD_MINIMUM.md`](MODEL_INPUTS_DASHBOARD_MINIMUM.md)
