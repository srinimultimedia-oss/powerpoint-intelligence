# Architecture decision register

## Confirmed by product owner
- AD-020 (1 November 2026): Owner approves one self-contained easySetup EXE covering x86/x64 Office, including the existing local template assets only for licensed computers. Implement a per-user bootstrapper around normal signed VSTO installation, automatic regeneration on signed Release builds and an optional source-edit watcher. Publisher certificate/trust and native installer acceptance remain evidence gates; no trust store or antivirus policy changes are authorized. See docs/implementation/easySetup-20261101.md.
- AD-019 (29 October 2026): Owner explicitly requests copying the installed template packages into the local plugin folder and updating the runtime reference. The local deployment uses apps/powerpoint-vsto/templates (62 byte-verified .potx packages, 497 slides) via the TemplateLibraryFolder add-in setting. This supersedes the installed-source-only reference for this computer. The local asset folder is Git-ignored; source package content and category names are preserved. Missing explicitly configured assets must report unavailability rather than silently switching sources.
- AD-018 (28 October 2026): Owner approves narrowly bounded antialias tolerance for template raster fidelity, alongside exact native formatting and geometry checks. Policy: identical dimensions; at most min(8, floor(pixel count × 0.00005)) changed pixels; RGB channel delta at most10/255; unchanged alpha; every changed pixel adjacent to an existing source edge. Do not blur, resize, mask regions, weaken native checks or mark offline replay as a new Office pass. See docs/implementation/template-fidelity-policy.md.
- AD-017 (26 October 2026): Owner approves live services in scope with separate configuration, and explicit Excel worksheet/range/column mapping. Missing, moved or structurally changed sources stop refresh; replacement requires reselecting and confirming the mapping. This resolves the moved-file behavior choice in Q-010; automatic rebinding is not permitted. Provider endpoints, identity, credential handling and production policies remain subject to Q-003/Q-004/Q-005. Implemented subsets still require native acceptance.
- AD-016 (26 October 2026): Owner rejects procedural layouts as substitutes for the think-cell source designs and approves hidden read-only source access for faithful template previews/import. Implement local external .potx source-library support, preserving source formatting and retaining the current destination presentation. Do not represent external-file reuse as an independently recreated editing engine or bundle third-party assets into product sources. Native rendered/save-reopen fidelity and any think-cell-dependent behavior require execution evidence.
- AD-015 (23 October 2026): Owner requests a first-release library of 500+ distinct editable templates and approves optional native builds and transitions. Implement original layouts; do not copy proprietary assets. The initial catalog has 520 procedural structural configurations across 26 families. Native Office editability, animation playback and save/reopen remain runtime acceptance gates; catalog count is not proof of those capabilities.
- AD-014 (22 October 2026): Owner requires chart commands to consume linked sources, avoid sample fallback, and generate reports within the current presentation. The local bounded table contracts are documented in `docs/implementation/data-source-links.md`; arbitrary field/range mapping and moved-file refresh policy remain open under Q-010. This change does not establish Office runtime acceptance.
- AD-013 (22 October 2026): Owner explicitly approves grouped editable native PowerPoint shapes for custom chart families, with canonical chart data retained. This resolves the native-chart-versus-shape-composition choice in Q-009/OD-06. Per-family capability, geometry fidelity, persistence, grouping/copy recovery and editability still require runtime evidence; this is not acceptance of an untested renderer.
- AD-012 (22 October 2026): The latest goal attachment restates the complete 98-requirement architecture/functional baseline and S0–S10 playbook. It supersedes the intervening generic technical blueprint as the active production objective. Windows/.NET Framework 4.8, native editability, on-prem services and Azure OpenAI boundaries remain fixed. Section 18 explicitly allows S1/S2 deterministic work in parallel with COM feasibility; this does not close S0 or promote demo functions to production acceptance. Open decisions remain open, and production evidence gates apply.
- AD-011 (22 October 2026): Owner requests a broad demonstration build from the pasted functional inventory before individual feature testing, superseding the previous S0-only implementation sequence for this demo increment. Owner explicitly chooses working local tools with AI/Tableau/MCP/enterprise connections marked not configured. This does not mark S0 stable, certify runtime behavior, or authorize simulated external results. Compile and structural checks continue; detailed acceptance testing is deferred at owner request.
- AD-001: Windows desktop PowerPoint first.
- AD-002: On-prem application hosting.
- AD-003: Azure OpenAI is the approved AI provider (external Azure dependency).
- AD-004: Native editable PowerPoint objects are the primary output model.
- AD-005: Full engineering documentation; Codex is expected to develop iteratively.
- AD-010 (22 October 2026): Owner requires configurable Windows desktop Office 2016 ("Office 16"), Office 2019 ("Office 19") and Microsoft 365 compatibility. Explicit profiles and configurable process architecture/minimum build are implemented; product-specific runtime acceptance is not yet verified. See `docs/implementation/office-compatibility.md`.

## Recommended, subject to proof-of-concept
- AD-006: VSTO/COM add-in built in C# targeting .NET Framework 4.8.
- AD-007: ASP.NET Core on-prem services; typed action orchestration.
- AD-008: TypeScript deterministic chart engine and tests.
- AD-009: Custom XML metadata plus stable object tags/IDs where supported.

## BLOCKED — obtain owner/environment confirmation
- Q-001 (partially resolved by AD-010): Office 2016, Office 2019 and Microsoft 365 families requested. Exact supported/tested builds, x86/x64 deployment matrix, update channels and licensing remain unresolved. Profile configuration is not runtime certification.
- Q-002: Windows desktop estate and software installation rights.
- Q-003: Azure OpenAI resource, region, deployment names, approved model versions and network egress/private endpoint topology.
- Q-004: On-prem identity provider, SSO protocols, user groups, service accounts.
- Q-005: Data classification, retention, prompts/logging and Azure processing policy.
- Q-006: Database platform, backup, HA/DR, infrastructure sizing and operational ownership.
- Q-007 (partially resolved by AD-020): easySetup EXE around signed VSTO deployment approved; approved distribution signing certificate and target-machine trust remain unresolved.
- Q-008: License model and offline grace policy.
- Q-009: Representation choice resolved by AD-013. Per-family capability/round-trip/editability acceptance remains an evidence gate.
- Q-010: Moved-file policy resolved by AD-017. Arbitrary range features, automatic updates, cross-element mappings and runtime acceptance remain incomplete.
- Q-011: Approved fonts, branding, accessibility and supported locales.
- Q-012: Exact performance baselines and reference test machines.

Do not block independent geometry tests on these decisions; do block production integration assumptions.
