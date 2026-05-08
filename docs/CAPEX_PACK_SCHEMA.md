# CAPEX pack schema (human-readable)

Canonical machine schema: [`schemas/project_capex_pack.v2026_01.schema.json`](../schemas/project_capex_pack.v2026_01.schema.json).

## Phase 1 (required)

- **`schema_version`**: must be `2026.01`.
- **`identity`**: `project_id`, `client_name`, `country`; optional `submitted_by`, `submitted_at`, `source_system`.
- **`capex_triplet`**:
  - `capex_ours` — total CAPEX for **our** solution (AI / company offering).
  - `capex_baseline_incumbent` — incumbent / legacy baseline CAPEX.
  - `capex_normal_led` — conventional LED retrofit benchmark CAPEX.
  - Optional: `currency`, `fx_rate_to_usd`, `inclusions_scope`, `confidence`, `source`.
- **`scale`**: `number_of_lights` (required); optional `number_of_poles`.

## Phase 2 (optional — unlocks full dashboard charts)

- **`opex_annual_by_scenario`**: keys `baseline`, `normal_led`, `emc`, `laas`; each may include `total_annual`, `electricity_annual`, `non_electric_annual`.
- **`maintenance_breakdown_by_scenario`**: same scenario keys; line items include `labor`, `materials`, `other`, `total_om`, `inspection`, `cleaning`, `testing`, `platform_software`, `spares`, `battery_reserve`.
- **`commercial_laas`**: `term_years`, `annual_service_fee`, `upfront_payment`, `escalation_pct_annual`, `currency`.
- **`calculated_kpis`**: mirrored payback / IRR / NPV from sales spreadsheet (validated separately by finance if used externally).
- **`attachments`**: `{ storage_uri, kind }[]` — pointers only.

## Chart mapping

| Chart | Required schema fragments |
|-------|---------------------------|
| Payback | Prefer `calculated_kpis.payback_years` + `commercial_laas`; crude fee-only heuristic available in mapper when IRR model not run |
| OPEX comparison | `opex_annual_by_scenario` |
| 运维费 breakdown | `maintenance_breakdown_by_scenario` |

## Examples

- Minimal: [`schemas/examples/minimal_valid.v2026_01.json`](../schemas/examples/minimal_valid.v2026_01.json)
- Phase 2 rich: [`schemas/examples/full_phase2_example.v2026_01.json`](../schemas/examples/full_phase2_example.v2026_01.json)
