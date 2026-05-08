# Runbook — questionnaire intake (operations)

## Dependencies

```bash
pip install -r requirements.txt
```

## Validate example payloads (CI)

```bash
pytest tests/test_questionnaire_schema.py -q
```

## Rotate Feishu app secret

1. Create new secret in Feishu developer console.
2. Update KMS / env vars on intake service (do **not** commit secrets).
3. Roll restart workers; verify webhook test event.

## OSS bucket hygiene

- Monthly review of lifecycle policies (raw uploads vs approved exports).
- Quarterly restore drill from RDS snapshot + spot-check attachment retrieval.

## Incident: validation failures spike

1. Check whether Feishu column rename broke mapping (`docs/feishu_field_mapping.md`).
2. Compare failing payload to [`schemas/examples/`](../schemas/examples/).
3. Bump `schema_version` only with [`schemas/CHANGELOG.md`](../schemas/CHANGELOG.md) update.
