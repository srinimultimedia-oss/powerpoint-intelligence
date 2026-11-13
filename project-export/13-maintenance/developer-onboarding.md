# Developer onboarding

1. Read export README, [master design](../15-deliverables/MASTER-TECHNICAL-DESIGN.md), module inventory and actual decision register `docs/DECISIONS.md:1`. Compare the exported commit/dirty-content baseline before relying on line numbers.
2. Use Windows desktop, Visual Studio Office/VSTO workload and .NET Framework 4.8 targeting pack for host work. Pure chart packages require Node (engine package declares >=24); locked AJV/TypeScript dependencies must already exist or be installed by the normal reviewed developer setup.
3. Understand three independent success levels: deterministic contracts/geometry tests, signed packaging integrity and actual interactive Office acceptance. A green portable suite does not prove VSTO registration, ribbon visibility or native save/reopen.
4. Locate UI routing in `ThisAddIn`, generated ribbon definitions/catalogs and context controller; find canonical codec/CustomXML attachment ownership before changing a feature. Select the existing compatible metadata namespace and version deliberately.
5. Run the appropriate existing test scripts in a developer session with presentations saved/closed for Office tests. This documentation task intentionally does not run scripts that would modify source/artifacts/configuration outside export.
6. Keep public display branding separate from internal COM IDs, namespaces, file paths and metadata identifiers. Preserve local library licenses and publisher trust rather than renaming identities without migration.

Build instructions and configuration are in [development setup](../11-build-deployment/development-setup.md). New production features need model validation, non-mutation rejection tests, native editing and save/reopen evidence. AI/service requirements must not be represented by fake returned content or endpoint-only readiness. `services/ai-gateway/README.md:1`.
