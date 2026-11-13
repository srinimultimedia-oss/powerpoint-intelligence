# Known issues and operational limits

The [security register](../09-security/security-gap-register.md), [technical debt](../14-audit/technical-debt-register.md), [architecture risks](../14-audit/architecture-risk-register.md) and [implementation gaps](../14-audit/implementation-gap-analysis.md) are the maintained issue catalogs for this baseline.

- Office16 API profiles are configurable, not product/build certificates. Native UI/COM cannot be accepted through a Linux or noninteractive token.
- PowerPoint can lose foreground while embedded Excel is active; context-property visibility is intentionally guarded and native UI tests can be BLOCKED rather than failed or passed.
- Linked Excel/chart mutation needs native-value observation before/after workbook close; RPC disconnection and busy-call failures are not resolved by assuming Close committed data.
- Missing/ambiguous CustomXMLPart association preserves native objects and requires explicit diagnosis. Save/reopen recovery cannot invent valid canonical state from arbitrary drawings.
- Scene/native adapter supports a bounded clustered rectangle profile; axes/legend omissions remain explicit diagnostics. V3 files are not compiled into active projects.
- Native mapped refresh rejects unsupported chart mixes/axes and source-shape changes. Source move/relink confirmation remains intentional.
- Procedural520-layout, original497-slide and sampled16-slide fidelity counts mean different things. Native transition/animation editing beyond tested samples is not certified.
- PIIDX digests do not authenticate a maliciously rebuilt database. Catalogs/previews can contain document text, metadata and paths; retention/permissions are still organizational requirements.
- Package fingerprint sorting differs across NLS/ICU hosts; normal Windows PowerShell was the validated historical packaging host.
- TestS1 dependency fallback and PIIDX evidence scope gaps require remediation separately; this export performs no repairs.
- Local-development signing trust is machine/user specific. Installing source-based versus extracted VSTO deployments can require normal migration/uninstall handling.
- AI/Azure/enterprise service requirements remain scaffolds or planning; no authenticated live connector is established by the UI profile editor.

Historical resolved development defects are retained in docs/artifacts and must not be reintroduced as current failures. Conversely older PASS logs must not hide current gaps. Use the exact module design and source references before modifying a boundary.
