# Architecture decision register

## Confirmed by product owner
- AD-001: Windows desktop PowerPoint first.
- AD-002: On-prem application hosting.
- AD-003: Azure OpenAI is the approved AI provider (external Azure dependency).
- AD-004: Native editable PowerPoint objects are the primary output model.
- AD-005: Full engineering documentation; Codex is expected to develop iteratively.

## Recommended, subject to proof-of-concept
- AD-006: VSTO/COM add-in built in C# targeting .NET Framework 4.8.
- AD-007: ASP.NET Core on-prem services; typed action orchestration.
- AD-008: TypeScript deterministic chart engine and tests.
- AD-009: Custom XML metadata plus stable object tags/IDs where supported.

## BLOCKED — obtain owner/environment confirmation
- Q-001: Supported Office versions, bitness, update channel and licensing.
- Q-002: Windows desktop estate and software installation rights.
- Q-003: Azure OpenAI resource, region, deployment names, approved model versions and network egress/private endpoint topology.
- Q-004: On-prem identity provider, SSO protocols, user groups, service accounts.
- Q-005: Data classification, retention, prompts/logging and Azure processing policy.
- Q-006: Database platform, backup, HA/DR, infrastructure sizing and operational ownership.
- Q-007: Distribution mechanism (MSI/ClickOnce/managed software deployment), signing certificate.
- Q-008: License model and offline grace policy.
- Q-009: Chart editability acceptance: native PowerPoint chart versus grouped native shapes for each chart family.
- Q-010: Excel link semantics and handling moved/renamed files.
- Q-011: Approved fonts, branding, accessibility and supported locales.
- Q-012: Exact performance baselines and reference test machines.

Do not block independent geometry tests on these decisions; do block production integration assumptions.
