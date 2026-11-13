# Documentation index

Start with [master design](15-deliverables/MASTER-TECHNICAL-DESIGN.md), [executive summary](EXECUTIVE-TECHNICAL-SUMMARY.md), [handover](15-deliverables/TECHNICAL-HANDOVER.md), [completeness](15-deliverables/DOCUMENTATION-COMPLETENESS-REPORT.md) and [manifest](EXPORT-MANIFEST.md).

CSV/JSON validation and traceability data are listed in the export manifest; authoring scripts/temporary binaries/gate machinery are local audit tooling, not product documentation or portable payload.

| Document | Purpose/title |
|---|---|
| [01-repository-analysis/dependency-inventory.md](01-repository-analysis/dependency-inventory.md) | Dependency inventory |
| [01-repository-analysis/entry-points.md](01-repository-analysis/entry-points.md) | Application and tooling entry points |
| [01-repository-analysis/repository-overview.md](01-repository-analysis/repository-overview.md) | Repository overview |
| [01-repository-analysis/source-analysis-coverage.md](01-repository-analysis/source-analysis-coverage.md) | Source analysis coverage |
| [01-repository-analysis/technology-stack.md](01-repository-analysis/technology-stack.md) | Verified technology stack |
| [02-architecture/application-architecture.md](02-architecture/application-architecture.md) | Application architecture |
| [02-architecture/architecture-decisions.md](02-architecture/architecture-decisions.md) | Implemented decisions and unresolved architecture |
| [02-architecture/architecture-diagrams.md](02-architecture/architecture-diagrams.md) | Architecture diagram index |
| [02-architecture/component-architecture.md](02-architecture/component-architecture.md) | Component architecture |
| [02-architecture/deployment-architecture.md](02-architecture/deployment-architecture.md) | Deployment architecture |
| [02-architecture/enterprise-architecture.md](02-architecture/enterprise-architecture.md) | Enterprise architecture |
| [02-architecture/solution-architecture.md](02-architecture/solution-architecture.md) | Solution and container architecture |
| [03-modules/individual-module-designs/core/catalog-query.md](03-modules/individual-module-designs/core/catalog-query.md) | PIIDX query models, search and navigation |
| [03-modules/individual-module-designs/core/catalog.md](03-modules/individual-module-designs/core/catalog.md) | Transactional PIIDX registration and refresh |
| [03-modules/individual-module-designs/core/chart-calculations.md](03-modules/individual-module-designs/core/chart-calculations.md) | Pure JavaScript and TypeScript numerical engines |
| [03-modules/individual-module-designs/core/domain-v2.md](03-modules/individual-module-designs/core/domain-v2.md) | Canonical chart v2 and explicit S0 migration |
| [03-modules/individual-module-designs/core/domain-v3.md](03-modules/individual-module-designs/core/domain-v3.md) | Authored chart v3 options and codec |
| [03-modules/individual-module-designs/core/library-config.md](03-modules/individual-module-designs/core/library-config.md) | Library profile and explicit legacy registration migration |
| [03-modules/individual-module-designs/core/library-refresh.md](03-modules/individual-module-designs/core/library-refresh.md) | Background filesystem refresh coordinator |
| [03-modules/individual-module-designs/core/native-renderer.md](03-modules/individual-module-designs/core/native-renderer.md) | Canonical projection and owned native shape renderer |
| [03-modules/individual-module-designs/core/path-policy.md](03-modules/individual-module-designs/core/path-policy.md) | Shared selected-input and generated-output policy |
| [03-modules/individual-module-designs/core/piidx-schema.md](03-modules/individual-module-designs/core/piidx-schema.md) | Versioned independent PIIDX schema |
| [03-modules/individual-module-designs/core/preview-cache.md](03-modules/individual-module-designs/core/preview-cache.md) | Validated PNG derivative cache |
| [03-modules/individual-module-designs/core/published-schemas.md](03-modules/individual-module-designs/core/published-schemas.md) | Published chart, scene and rendered-record schemas |
| [03-modules/individual-module-designs/core/scene-contract.md](03-modules/individual-module-designs/core/scene-contract.md) | Scene DTOs and cross-language hash boundary |
| [03-modules/individual-module-designs/core/source-extraction.md](03-modules/individual-module-designs/core/source-extraction.md) | Read-only PPTX/POTX metadata extraction |
| [03-modules/individual-module-designs/core/sqlite.md](03-modules/individual-module-designs/core/sqlite.md) | Windows native SQLite bridge |
| [03-modules/individual-module-designs/core/strict-json.md](03-modules/individual-module-designs/core/strict-json.md) | Bounded JSON and Unicode utilities |
| [03-modules/individual-module-designs/core/svg-preview.md](03-modules/individual-module-designs/core/svg-preview.md) | Internal SVG geometry inspection |
| [03-modules/individual-module-designs/core/ts-scene.md](03-modules/individual-module-designs/core/ts-scene.md) | TypeScript scene construction and v2 file adapter |
| [03-modules/individual-module-designs/operations/artifact-maintenance.md](03-modules/individual-module-designs/operations/artifact-maintenance.md) | OPS-09 — Artifact maintenance and repository dependency audit |
| [03-modules/individual-module-designs/operations/build-provenance.md](03-modules/individual-module-designs/operations/build-provenance.md) | OPS-05 — Build provenance, recovery and dependency restoration |
| [03-modules/individual-module-designs/operations/com-automation.md](03-modules/individual-module-designs/operations/com-automation.md) | OPS-02 — COM automation interface |
| [03-modules/individual-module-designs/operations/deployment-bootstrap.md](03-modules/individual-module-designs/operations/deployment-bootstrap.md) | OPS-04 — Per-user deployment bootstrapper |
| [03-modules/individual-module-designs/operations/diagnostics.md](03-modules/individual-module-designs/operations/diagnostics.md) | OPS-06 — Activation and scanner diagnostics |
| [03-modules/individual-module-designs/operations/offline-alignment-repair.md](03-modules/individual-module-designs/operations/offline-alignment-repair.md) | OPS-08 — Offline alignment repair utility |
| [03-modules/individual-module-designs/operations/packaging-build.md](03-modules/individual-module-designs/operations/packaging-build.md) | OPS-03 — Signed packaging build and update watcher |
| [03-modules/individual-module-designs/operations/service-profiles.md](03-modules/individual-module-designs/operations/service-profiles.md) | OPS-01 — Service profile configuration and AI scaffold |
| [03-modules/individual-module-designs/operations/validation-orchestration.md](03-modules/individual-module-designs/operations/validation-orchestration.md) | OPS-07 — Validation orchestration and CI |
| [03-modules/individual-module-designs/vsto/agendas.md](03-modules/individual-module-designs/vsto/agendas.md) | Managed chapters/contents and legacy presentation agenda |
| [03-modules/individual-module-designs/vsto/commands-ribbon.md](03-modules/individual-module-designs/vsto/commands-ribbon.md) | Ribbon catalog, dispatch, icons, browser and automation facade |
| [03-modules/individual-module-designs/vsto/connectors.md](03-modules/individual-module-designs/vsto/connectors.md) | Universal connector geometry and endpoint bindings |
| [03-modules/individual-module-designs/vsto/context-properties.md](03-modules/individual-module-designs/vsto/context-properties.md) | Selection-aware floating property box |
| [03-modules/individual-module-designs/vsto/indicators.md](03-modules/individual-module-designs/vsto/indicators.md) | Harvey-ball and checkbox editable vectors |
| [03-modules/individual-module-designs/vsto/library-gallery.md](03-modules/individual-module-designs/vsto/library-gallery.md) | Thumbnail gallery and PIIDX workspace UI |
| [03-modules/individual-module-designs/vsto/lifecycle-host.md](03-modules/individual-module-designs/vsto/lifecycle-host.md) | VSTO lifecycle, compatibility and S0 native host |
| [03-modules/individual-module-designs/vsto/linked-data.md](03-modules/individual-module-designs/vsto/linked-data.md) | Linked files, Excel mapping and source manager |
| [03-modules/individual-module-designs/vsto/managed-text.md](03-modules/individual-module-designs/vsto/managed-text.md) | Text box, pentagon, chevron and rounded native elements |
| [03-modules/individual-module-designs/vsto/native-chart-controls.md](03-modules/individual-module-designs/vsto/native-chart-controls.md) | Native axes, series and annotations |
| [03-modules/individual-module-designs/vsto/native-formatting.md](03-modules/individual-module-designs/vsto/native-formatting.md) | Native formatting plans and restoration transactions |
| [03-modules/individual-module-designs/vsto/process-grid.md](03-modules/individual-module-designs/vsto/process-grid.md) | Process-grid layout and retained structural history |
| [03-modules/individual-module-designs/vsto/productivity.md](03-modules/individual-module-designs/vsto/productivity.md) | Local symbol, decimal, picture-copy and chart workbook tools |
| [03-modules/individual-module-designs/vsto/quality.md](03-modules/individual-module-designs/vsto/quality.md) | Read-only presentation heuristics and navigation |
| [03-modules/individual-module-designs/vsto/renderer-bridge.md](03-modules/individual-module-designs/vsto/renderer-bridge.md) | Native scene-renderer ribbon adapter |
| [03-modules/individual-module-designs/vsto/service-config.md](03-modules/individual-module-designs/vsto/service-config.md) | Nonsecret external-service configuration UI |
| [03-modules/individual-module-designs/vsto/template-catalog.md](03-modules/individual-module-designs/vsto/template-catalog.md) | Designed template layouts, editable fields and palettes |
| [03-modules/individual-module-designs/vsto/template-insertion.md](03-modules/individual-module-designs/vsto/template-insertion.md) | Template insertion, source-original sessions and layout links |
| [03-modules/individual-module-designs/vsto/xml-attachments.md](03-modules/individual-module-designs/vsto/xml-attachments.md) | Canonical XML ownership and compensating replacement |
| [03-modules/module-dependency-matrix.md](03-modules/module-dependency-matrix.md) | Module dependency matrix |
| [03-modules/module-inventory.md](03-modules/module-inventory.md) | Module inventory |
| [04-powerpoint-integration/addin-lifecycle.md](04-powerpoint-integration/addin-lifecycle.md) | Add-in lifecycle and event ownership |
| [04-powerpoint-integration/analysis-method.md](04-powerpoint-integration/analysis-method.md) | VSTO analysis method and explicit coverage boundary |
| [04-powerpoint-integration/interoperability.md](04-powerpoint-integration/interoperability.md) | Interoperability, threading and process boundaries |
| [04-powerpoint-integration/native-editing-engine.md](04-powerpoint-integration/native-editing-engine.md) | Native editing and persistence mechanisms |
| [04-powerpoint-integration/powerpoint-integration-overview.md](04-powerpoint-integration/powerpoint-integration-overview.md) | PowerPoint integration overview |
| [04-powerpoint-integration/rendering-engine.md](04-powerpoint-integration/rendering-engine.md) | Rendering engines and the ribbon projection |
| [04-powerpoint-integration/ribbon-architecture.md](04-powerpoint-integration/ribbon-architecture.md) | Ribbon architecture and command routing |
| [04-powerpoint-integration/shape-and-slide-operations.md](04-powerpoint-integration/shape-and-slide-operations.md) | Shape, text, table and slide operation design |
| [05-features/complete-feature-inventory.md](05-features/complete-feature-inventory.md) | Complete NexWave feature inventory |
| [05-features/feature-implementation-details.md](05-features/feature-implementation-details.md) | Feature implementation details |
| [05-features/feature-status-matrix.md](05-features/feature-status-matrix.md) | Feature status matrix |
| [05-features/feature-workflows.md](05-features/feature-workflows.md) | Feature execution workflows |
| [06-ai-architecture/ai-data-flow.md](06-ai-architecture/ai-data-flow.md) | AI data flow and context boundary |
| [06-ai-architecture/ai-security.md](06-ai-architecture/ai-security.md) | AI security |
| [06-ai-architecture/ai-solution-architecture.md](06-ai-architecture/ai-solution-architecture.md) | AI solution architecture |
| [06-ai-architecture/ai-workflows.md](06-ai-architecture/ai-workflows.md) | Implemented AI-adjacent workflow |
| [06-ai-architecture/azure-ai-integration.md](06-ai-architecture/azure-ai-integration.md) | Azure AI integration |
| [06-ai-architecture/prompt-orchestration.md](06-ai-architecture/prompt-orchestration.md) | Prompt and action orchestration |
| [07-api-design/api-inventory.md](07-api-design/api-inventory.md) | API and integration inventory |
| [07-api-design/external-interfaces.md](07-api-design/external-interfaces.md) | External interfaces |
| [07-api-design/internal-interfaces.md](07-api-design/internal-interfaces.md) | Internal interfaces |
| [07-api-design/native-event-declarations.md](07-api-design/native-event-declarations.md) | Native and event contract declarations |
| [07-api-design/openapi/README.md](07-api-design/openapi/README.md) | OpenAPI applicability |
| [07-api-design/request-response-contracts.md](07-api-design/request-response-contracts.md) | Request/response and file contracts |
| [08-data-design/class-diagrams.md](08-data-design/class-diagrams.md) | Class and relationship diagrams |
| [08-data-design/data-dictionary.md](08-data-design/data-dictionary.md) | Data dictionary |
| [08-data-design/data-models.md](08-data-design/data-models.md) | Data models |
| [08-data-design/persistence-design.md](08-data-design/persistence-design.md) | Persistence design |
| [08-data-design/serialization-contracts.md](08-data-design/serialization-contracts.md) | Serialization and compatibility contracts |
| [09-security/authentication-authorization.md](09-security/authentication-authorization.md) | Authentication and authorization |
| [09-security/secrets-management.md](09-security/secrets-management.md) | Secrets management and sensitive data |
| [09-security/security-architecture.md](09-security/security-architecture.md) | Security architecture |
| [09-security/security-gap-register.md](09-security/security-gap-register.md) | Security gap register |
| [09-security/threat-model.md](09-security/threat-model.md) | Threat model |
| [10-workflows/activity-diagrams.md](10-workflows/activity-diagrams.md) | Activity diagrams |
| [10-workflows/data-flow-diagrams.md](10-workflows/data-flow-diagrams.md) | Data-flow and trust-boundary diagrams |
| [10-workflows/error-handling-flows.md](10-workflows/error-handling-flows.md) | Error handling and recovery flows |
| [10-workflows/sequence-diagrams.md](10-workflows/sequence-diagrams.md) | Execution sequence diagrams |
| [10-workflows/workflow-inventory.md](10-workflows/workflow-inventory.md) | Workflow inventory |
| [11-build-deployment/build-instructions.md](11-build-deployment/build-instructions.md) | Build instructions |
| [11-build-deployment/configuration-reference.md](11-build-deployment/configuration-reference.md) | Configuration reference |
| [11-build-deployment/deployment-guide.md](11-build-deployment/deployment-guide.md) | Deployment guide and acceptance |
| [11-build-deployment/development-setup.md](11-build-deployment/development-setup.md) | Development setup |
| [11-build-deployment/packaging-installation.md](11-build-deployment/packaging-installation.md) | Packaging, installation and upgrades |
| [11-build-deployment/troubleshooting-guide.md](11-build-deployment/troubleshooting-guide.md) | Troubleshooting |
| [12-testing/coverage-gap-analysis.md](12-testing/coverage-gap-analysis.md) | Test coverage gaps |
| [12-testing/test-architecture.md](12-testing/test-architecture.md) | Testing architecture |
| [12-testing/test-inventory.md](12-testing/test-inventory.md) | Complete authored test inventory |
| [12-testing/test-results.md](12-testing/test-results.md) | Current validation and historical evidence |
| [13-maintenance/debugging-guide.md](13-maintenance/debugging-guide.md) | Debugging guide |
| [13-maintenance/developer-onboarding.md](13-maintenance/developer-onboarding.md) | Developer onboarding |
| [13-maintenance/extension-guide.md](13-maintenance/extension-guide.md) | Extension guide |
| [13-maintenance/known-issues.md](13-maintenance/known-issues.md) | Known issues and operational limits |
| [13-maintenance/operational-runbook.md](13-maintenance/operational-runbook.md) | Operational runbook |
| [14-audit/architecture-risk-register.md](14-audit/architecture-risk-register.md) | Architecture risk register |
| [14-audit/assumptions-and-unknowns.md](14-audit/assumptions-and-unknowns.md) | Assumptions and unresolved questions |
| [14-audit/documentation-validation.md](14-audit/documentation-validation.md) | Documentation validation method and results |
| [14-audit/implementation-gap-analysis.md](14-audit/implementation-gap-analysis.md) | Implementation gap analysis |
| [14-audit/technical-debt-register.md](14-audit/technical-debt-register.md) | Technical debt register |
| [15-deliverables/DOCUMENTATION-COMPLETENESS-REPORT.md](15-deliverables/DOCUMENTATION-COMPLETENESS-REPORT.md) | Documentation completeness and acceptance |
| [15-deliverables/MASTER-TECHNICAL-DESIGN.md](15-deliverables/MASTER-TECHNICAL-DESIGN.md) | NexWave — Master Technical Design |
| [15-deliverables/TECHNICAL-HANDOVER.md](15-deliverables/TECHNICAL-HANDOVER.md) | Technical handover |
| [EXECUTIVE-TECHNICAL-SUMMARY.md](EXECUTIVE-TECHNICAL-SUMMARY.md) | Executive technical summary |
| [EXPORT-MANIFEST.md](EXPORT-MANIFEST.md) | Export manifest |
| [README.md](README.md) | NexWave technical design export |
