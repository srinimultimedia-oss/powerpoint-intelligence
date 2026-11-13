# Module inventory

**VERIFIED source designs:** coherent modules are counted by design document, not class or directory count. A source file may serve several related modules; the source mapping preserves that many-to-many relationship. Each design addresses the17 required dimensions. Runtime evidence and unsupported paths are explicitly separate.

| ID | Module | Design | Sources |
|---|---|---|---|
| VSTO-01 | VSTO lifecycle, compatibility and S0 native host | [vsto/lifecycle-host](individual-module-designs/vsto/lifecycle-host.md) | 9 |
| VSTO-02 | Ribbon catalog, dispatch, icons, browser and automation facade | [vsto/commands-ribbon](individual-module-designs/vsto/commands-ribbon.md) | 7 |
| VSTO-03 | Linked files, Excel mapping and source manager | [vsto/linked-data](individual-module-designs/vsto/linked-data.md) | 5 |
| VSTO-04 | Native axes, series and annotations | [vsto/native-chart-controls](individual-module-designs/vsto/native-chart-controls.md) | 3 |
| VSTO-05 | Selection-aware floating property box | [vsto/context-properties](individual-module-designs/vsto/context-properties.md) | 7 |
| VSTO-06 | Native formatting plans and restoration transactions | [vsto/native-formatting](individual-module-designs/vsto/native-formatting.md) | 1 |
| VSTO-07 | Text box, pentagon, chevron and rounded native elements | [vsto/managed-text](individual-module-designs/vsto/managed-text.md) | 2 |
| VSTO-08 | Process-grid layout and retained structural history | [vsto/process-grid](individual-module-designs/vsto/process-grid.md) | 1 |
| VSTO-09 | Harvey-ball and checkbox editable vectors | [vsto/indicators](individual-module-designs/vsto/indicators.md) | 1 |
| VSTO-10 | Universal connector geometry and endpoint bindings | [vsto/connectors](individual-module-designs/vsto/connectors.md) | 1 |
| VSTO-11 | Managed chapters/contents and legacy presentation agenda | [vsto/agendas](individual-module-designs/vsto/agendas.md) | 3 |
| VSTO-12 | Designed template layouts, editable fields and palettes | [vsto/template-catalog](individual-module-designs/vsto/template-catalog.md) | 1 |
| VSTO-13 | Template insertion, source-original sessions and layout links | [vsto/template-insertion](individual-module-designs/vsto/template-insertion.md) | 1 |
| VSTO-14 | Thumbnail gallery and PIIDX workspace UI | [vsto/library-gallery](individual-module-designs/vsto/library-gallery.md) | 2 |
| VSTO-15 | Local symbol, decimal, picture-copy and chart workbook tools | [vsto/productivity](individual-module-designs/vsto/productivity.md) | 1 |
| VSTO-16 | Read-only presentation heuristics and navigation | [vsto/quality](individual-module-designs/vsto/quality.md) | 1 |
| VSTO-17 | Canonical XML ownership and compensating replacement | [vsto/xml-attachments](individual-module-designs/vsto/xml-attachments.md) | 1 |
| VSTO-18 | Nonsecret external-service configuration UI | [vsto/service-config](individual-module-designs/vsto/service-config.md) | 1 |
| VSTO-19 | Native scene-renderer ribbon adapter | [vsto/renderer-bridge](individual-module-designs/vsto/renderer-bridge.md) | 1 |
| domain-v2 | Canonical chart v2 and explicit S0 migration | [core/domain-v2](individual-module-designs/core/domain-v2.md) | 4 |
| domain-v3 | Authored chart v3 options and codec | [core/domain-v3](individual-module-designs/core/domain-v3.md) | 2 |
| strict-json | Bounded JSON and Unicode utilities | [core/strict-json](individual-module-designs/core/strict-json.md) | 1 |
| scene-contract | Scene DTOs and cross-language hash boundary | [core/scene-contract](individual-module-designs/core/scene-contract.md) | 3 |
| native-renderer | Canonical projection and owned native shape renderer | [core/native-renderer](individual-module-designs/core/native-renderer.md) | 2 |
| sqlite | Windows native SQLite bridge | [core/sqlite](individual-module-designs/core/sqlite.md) | 1 |
| piidx-schema | Versioned independent PIIDX schema | [core/piidx-schema](individual-module-designs/core/piidx-schema.md) | 1 |
| source-extraction | Read-only PPTX/POTX metadata extraction | [core/source-extraction](individual-module-designs/core/source-extraction.md) | 1 |
| catalog | Transactional PIIDX registration and refresh | [core/catalog](individual-module-designs/core/catalog.md) | 1 |
| catalog-query | PIIDX query models, search and navigation | [core/catalog-query](individual-module-designs/core/catalog-query.md) | 3 |
| preview-cache | Validated PNG derivative cache | [core/preview-cache](individual-module-designs/core/preview-cache.md) | 1 |
| library-config | Library profile and explicit legacy registration migration | [core/library-config](individual-module-designs/core/library-config.md) | 1 |
| library-refresh | Background filesystem refresh coordinator | [core/library-refresh](individual-module-designs/core/library-refresh.md) | 1 |
| path-policy | Shared selected-input and generated-output policy | [core/path-policy](individual-module-designs/core/path-policy.md) | 1 |
| chart-calculations | Pure JavaScript and TypeScript numerical engines | [core/chart-calculations](individual-module-designs/core/chart-calculations.md) | 5 |
| ts-scene | TypeScript scene construction and v2 file adapter | [core/ts-scene](individual-module-designs/core/ts-scene.md) | 6 |
| svg-preview | Internal SVG geometry inspection | [core/svg-preview](individual-module-designs/core/svg-preview.md) | 2 |
| published-schemas | Published chart, scene and rendered-record schemas | [core/published-schemas](individual-module-designs/core/published-schemas.md) | 3 |
| OPS-01 | service-profiles | [operations/service-profiles](individual-module-designs/operations/service-profiles.md) | 4 |
| OPS-02 | com-automation | [operations/com-automation](individual-module-designs/operations/com-automation.md) | 2 |
| OPS-03 | packaging-build | [operations/packaging-build](individual-module-designs/operations/packaging-build.md) | 12 |
| OPS-04 | deployment-bootstrap | [operations/deployment-bootstrap](individual-module-designs/operations/deployment-bootstrap.md) | 1 |
| OPS-05 | build-provenance | [operations/build-provenance](individual-module-designs/operations/build-provenance.md) | 13 |
| OPS-06 | diagnostics | [operations/diagnostics](individual-module-designs/operations/diagnostics.md) | 4 |
| OPS-07 | validation-orchestration | [operations/validation-orchestration](individual-module-designs/operations/validation-orchestration.md) | 27 |
| OPS-08 | offline-alignment-repair | [operations/offline-alignment-repair](individual-module-designs/operations/offline-alignment-repair.md) | 3 |
| OPS-09 | artifact-maintenance | [operations/artifact-maintenance](individual-module-designs/operations/artifact-maintenance.md) | 4 |

The core map includes authored-but-uncompiled V3 and published schemas rather than hiding them. Operations designs include build/test/diagnostic tooling separately from runtime modules. No web backend is invented. See [source mapping](source-module-map.csv).
