# Mainland China–friendly storage architecture (questionnaire + artifacts)

## Roles

| Layer | Recommended (examples) | Holds |
|-------|-------------------------|-------|
| **Object storage** | Alibaba OSS, Tencent COS, Volcengine TOS | PDF/PNG/XLSX attachments, exported dashboards |
| **Relational DB** | Alibaba RDS MySQL / PolarDB, Tencent TencentDB | Project metadata, submission versions, validation status |
| **Secrets** | Cloud KMS (same vendor as OSS/DB) | Feishu app secrets, OSS keys |
| **Compute** | ECS / Cloud Functions / Container Service | Intake API, agent workers |

## Why not GitHub for filled questionnaires?

- **Availability** from mainland networks can be unreliable without mirrors/VPN.
- **Access control** for client pricing and commercial terms is weaker in git.
- **Binary blobs** (photos, scans) bloat repositories and violate least-privilege.

**GitHub remains appropriate for:** JSON Schema, validators, dashboard code, infra-as-code templates — **not** for production PII/financial submissions.

## Backup and DR (within China)

- Enable **OSS versioning** + cross-region replication **inside China** where policy allows (e.g. Shanghai ↔ Shenzhen).
- RDS **automated snapshots** + periodic restore drills documented in `RUNBOOK.md`.
- **Retention**: attach lifecycle rules for raw uploads vs approved exports.

## Firewall-friendly access

- Primary endpoints on **domestic** domains/CDN.
- Overseas staff: **approved** VPN or read-only export bucket — avoid storing authoritative data only on GitHub.
