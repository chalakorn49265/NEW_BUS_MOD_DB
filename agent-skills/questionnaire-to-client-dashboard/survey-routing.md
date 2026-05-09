# Survey routing: which dashboard story?

Use **`#` IDs** from `questionnaire_01_input.xlsx`. Human-readable labels live in `questionnaire_01_zh_map.py` (**ZH_MAP**) in the reference repo.

## Block J — incumbent EMC / ESCO / third-party contract

| ID | Route when relevant |
|----|---------------------|
| **J1** | **全覆盖 / 仅部分** → incumbent EMC exists; surface contract layer in UI copy, constrain “before” baseline and who pays electricity (**J4**). |
| **J2** | Remaining contract term (years)—timeline for coexistence with new structure. |
| **J3** | Mechanic: 节能分享 / 托管 / 租赁 / 混合 — affects how savings vs fees are narrated. |
| **J4** | **Who pays 市电** under current contract — critical for bill attribution. |
| **J5–J6** | Constraints and rough annual cash tied to arrangement—risk/disclosure section. |

**Dashboard emphasis:** Retrofit / savings / fee-split economics; clarify overlap with existing EMC (avoid double-counting baseline savings).

## Blocks H / I / K — ownership, term, LaaS-style structure

| ID | Route when relevant |
|----|---------------------|
| **H7** | Asset registration requirement — affects lease vs ownership narrative. |
| **H10** | Acceptable contract term — align scenario length. |
| **I2** | Usership vs ownership — **LaaS**-style acceptance. |
| **I5** | End-of-term disposal — exit scenarios. |
| **K6** | Case study / data consent — optional footer disclaimer. |

**Dashboard emphasis:** **LaaS** IRR pages, feasible envelope, subscription vs baseline comparisons.

## Block G — payer and budget stress

| ID | Route |
|----|--------|
| **G1** | Who pays — informs stakeholder labels in KPIs. |
| **G11–G12** | Budget stress / evidence — optional risk callout. |

## Energy & money reconciliation (always)

- **B1** vs **B2+B3+B4** sanity check (Definitions in master workbook).
- **C4** + **C4a–C4e** for TOU vs flat — drives tariff inputs.

## Quick decision tree

1. **J1** = 无 / 未知 → standard retrofit baseline narrative.
2. **J1** = 全覆盖 / 仅部分 → **EMC incumbent** path + **J2–J6** in sidebar or expanders.
3. **I2** / **H10** strongly toward lease-like structure → route to **LaaS / feasible envelope** pattern dashboards.
4. Otherwise → generic client economics + tier comparison as needed.
