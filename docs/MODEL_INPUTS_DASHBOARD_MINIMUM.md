# Dashboard-minimum model inputs (distributable)

This document lists **only** the inputs needed to build the same **dashboard families** supported by [`schemas/project_capex_pack.v2026_01.schema.json`](../schemas/project_capex_pack.v2026_01.schema.json):

- CAPEX triplet comparison  
- Payback (via `commercial_laas` + optional mirrored KPIs)  
- Annual **OPEX** comparison: `baseline` → `normal_led` → `emc` → **`laas`**  
- **运维费** (maintenance / O&M) breakdown **per scenario**

It **does not** include full IPP-style project finance (debt schedule, WC, terminal value, detailed tax). Those belong to a separate “full model” workbook.

**Companion machine-readable template:** [`csv/model_inputs_dashboard_minimum.csv`](csv/model_inputs_dashboard_minimum.csv) (wide format: one row per project; header = flattened JSON paths).

---

## How this maps to JSON

Paste row values into a nested object:

- Column names use **dot paths** (e.g. `capex_triplet.capex_ours`) matching JSON keys under `project_capex_pack`.
- Set top-level `schema_version` to **`2026.01`**.
- **Currency:** all money amounts in the CSV row share `capex_triplet.currency` unless you also fill `commercial_laas.currency` when fee currency differs (rare). Use `capex_triplet.fx_rate_to_usd` when amounts are in local currency.

Validate with:

```python
from questionnaire.validate_payload import validate_project_capex_pack
import json

with open("pack.json") as f:
    validate_project_capex_pack(json.load(f))
```

---

## Phase 1 — Required for CAPEX + scale

| Field key | English label | 中文提示 | Type | Unit | Required | Notes |
|-----------|---------------|---------|------|------|----------|--------|
| `schema_version` | Schema version | 数据版本 | Quant | string | Y | Must be `2026.01` |
| `identity.project_id` | Project ID | 项目编号 | Qual | text | Y | Unique per submission |
| `identity.client_name` | Client name | 客户名称 | Qual | text | Y | |
| `identity.country` | Country | 国家/地区 | Qual | text | Y | ISO country name or CN |
| `identity.submitted_by` | Submitted by | 填写人 | Qual | text | N | |
| `identity.submitted_at` | Submitted at | 提交时间 | Quant | ISO 8601 | N | |
| `identity.sales_region` | Sales region | 销售大区 | Qual | text | N | Internal |
| `identity.source_system` | Source system | 来源 | Qual | text | N | e.g. `csv_import_v1` |
| `capex_triplet.capex_ours` | Total CAPEX — our solution | 我方方案总投资 | Quant | currency | Y | AI / company offering |
| `capex_triplet.capex_baseline_incumbent` | Total CAPEX — incumbent baseline | 现状/基准总投资 | Quant | currency | Y | Legacy / like-for-like baseline |
| `capex_triplet.capex_normal_led` | Total CAPEX — normal LED benchmark | 普通LED方案总投资 | Quant | currency | Y | Conventional LED retrofit |
| `capex_triplet.currency` | Currency for CAPEX amounts | 币种 | Qual | ISO 4217 | N | Default USD |
| `capex_triplet.fx_rate_to_usd` | FX rate to USD | 对美元汇率 | Quant | ratio | N | If currency ≠ USD |
| `capex_triplet.inclusions_scope` | CAPEX scope / inclusions | 投资范围说明 | Qual | text | N | Fixtures, civil, controls — keep comparable across three lines |
| `capex_triplet.confidence` | Estimate confidence | 置信度 | Qual | enum | N | `high` / `medium` / `low` |
| `capex_triplet.source` | Data source | 数据来源 | Qual | enum | N | `client_quote` / `internal_estimate` / `benchmark_table` / `other` |
| `scale.number_of_lights` | Number of luminaires | 灯具数量 | Quant | count | Y | For per-light charts |
| `scale.number_of_poles` | Number of poles | 灯杆数量 | Quant | count | N | If different from lights |

---

## Phase 2 — Required for non-empty OPEX + 运维费 charts

Fill **all four scenarios** with the **same definitions** of “baseline”, “normal LED”, “EMC”, “LaaS”.

### Annual OPEX by scenario

| Field key | English label | 中文提示 | Type | Unit | Required for OPEX chart |
|-----------|---------------|---------|------|------|-------------------------|
| `opex_annual_by_scenario.baseline.total_annual` | Baseline — total annual OPEX | 现状年运维总成本 | Quant | currency | Y |
| `opex_annual_by_scenario.baseline.electricity_annual` | Baseline — annual electricity cost | 现状年电费 | Quant | currency | N | Split if known |
| `opex_annual_by_scenario.baseline.non_electric_annual` | Baseline — annual non-electric OPEX | 现状年非电费运维 | Quant | currency | N | |
| `opex_annual_by_scenario.normal_led.total_annual` | Normal LED — total annual OPEX | 普通LED年运维总成本 | Quant | currency | Y |
| `opex_annual_by_scenario.normal_led.electricity_annual` | Normal LED — annual electricity | 普通LED年电费 | Quant | currency | N | |
| `opex_annual_by_scenario.normal_led.non_electric_annual` | Normal LED — non-electric OPEX | 普通LED年非电费运维 | Quant | currency | N | |
| `opex_annual_by_scenario.emc.total_annual` | EMC — total annual OPEX | EMC年运维总成本 | Quant | currency | Y |
| `opex_annual_by_scenario.emc.electricity_annual` | EMC — annual electricity | EMC年电费 | Quant | currency | N | |
| `opex_annual_by_scenario.emc.non_electric_annual` | EMC — non-electric OPEX | EMC年非电费运维 | Quant | currency | N | |
| `opex_annual_by_scenario.laas.total_annual` | LaaS — total annual client cash cost | LaaS模式年总支出 | Quant | currency | Y | Typically subscription + residual |
| `opex_annual_by_scenario.laas.electricity_annual` | LaaS — annual electricity | LaaS年电费 | Quant | currency | N | Often near 0 if solar/off-grid story |
| `opex_annual_by_scenario.laas.non_electric_annual` | LaaS — non-electric OPEX | LaaS年非电费运维 | Quant | currency | N | |

### 运维费 breakdown by scenario (same keys each scenario)

| Field key suffix | English label | 中文提示 | Type | Unit |
|------------------|---------------|---------|------|------|
| `.labor` | Labor | 人工费 | Quant | currency |
| `.materials` | Materials | 材料费 | Quant | currency |
| `.other` | Other | 其他 | Quant | currency |
| `.total_om` | Total O&M (if reported separately) | 运维费合计 | Quant | currency |
| `.inspection` | Inspection | 巡检费 | Quant | currency |
| `.cleaning` | Cleaning | 清洗费 | Quant | currency |
| `.testing` | Testing | 检测费 | Quant | currency |
| `.platform_software` | Platform / software | 平台/软件 | Quant | currency |
| `.spares` | Spares | 备件 | Quant | currency |
| `.battery_reserve` | Battery reserve | 电池/储备 | Quant | currency |

Prefixes (repeat for each scenario):

- `maintenance_breakdown_by_scenario.baseline.*`
- `maintenance_breakdown_by_scenario.normal_led.*`
- `maintenance_breakdown_by_scenario.emc.*`
- `maintenance_breakdown_by_scenario.laas.*`

Optional per-block: `maintenance_breakdown_by_scenario.{scenario}.currency` if mixed currencies (avoid if possible).

---

## LaaS commercial (payback / subscription)

| Field key | English label | 中文提示 | Type | Unit | Required |
|-----------|---------------|---------|------|------|----------|
| `commercial_laas.term_years` | LaaS term | LaaS服务期限（年） | Quant | years | N |
| `commercial_laas.annual_service_fee` | Annual service fee | 年度服务费 | Quant | currency | N |
| `commercial_laas.upfront_payment` | Upfront payment | 首付款 | Quant | currency | N |
| `commercial_laas.escalation_pct_annual` | Annual fee escalation | 年费递增比例 | Quant | decimal | N | e.g. `0.03` for 3% |
| `commercial_laas.currency` | Fee currency | 服务费币种 | Qual | ISO 4217 | N |

---

## Optional KPI mirror (from spreadsheet / finance)

| Field key | English label | Type | Unit |
|-----------|---------------|------|------|
| `calculated_kpis.payback_years` | Payback period | Quant | years |
| `calculated_kpis.irr_annual` | IRR (annual) | Quant | decimal |
| `calculated_kpis.npv` | NPV | Quant | currency |
| `calculated_kpis.notes` | KPI notes | Qual | text |

---

## Messaging (client-facing guardrails)

| Field key | English label | Type |
|-----------|---------------|------|
| `messaging_constraints.do_not_show_client` | Do not show client | Qual |
| `messaging_constraints.preferred_currency_display` | Preferred display currency | Qual |

---

## CSV import / export

1. Use **[`csv/model_inputs_dashboard_minimum.csv`](csv/model_inputs_dashboard_minimum.csv)** as the header row for Google Sheets or Feishu Base column names (exact match recommended).
2. Each **data row** = one project. Do not merge cells.
3. Empty optional cells are allowed; required Phase 1 cells must be filled before JSON validation.
4. After editing, convert row → nested JSON (tooling can be added later) and run `validate_project_capex_pack`.

The second row in the CSV file is an **illustrative example** only — replace with real project data.
