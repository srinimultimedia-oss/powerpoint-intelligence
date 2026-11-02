# PowerPoint Intelligence — Feature Requirements Index

**Status:** Requirements index for implementation planning. No feature listed here is implied to be implemented or tested.

**Authoritative specification:** [`architecture-functional-design.docx`](architecture-functional-design.docx) (original 98-feature requirements, acceptance tests, architecture and engineering decisions). Read that document for the detailed specification. This Markdown file is a navigation aid, not a replacement for the design document.

## Product constraints

- **Target:** Windows desktop Microsoft PowerPoint first.
- **PowerPoint host:** C# VSTO / COM add-in using a compatible .NET Framework runtime.
- **Services:** On-premises application services; Azure OpenAI is an external Azure-hosted inference service accessed through a controlled gateway.
- **Output:** Editable native PowerPoint objects, with canonical versioned chart/document metadata.
- **Engineering:** Independent original implementation; do not copy think-cell proprietary assets or code.
- **AI:** Model outputs are proposals expressed as validated, typed action plans. Deterministic engines perform calculations and PowerPoint changes, subject to user authorization and approval.

## Functional requirements — core (68)

### Chart families (CHT-001–CHT-013)

| ID | Feature |
|---|---|
| CHT-001 | Clustered bar and column charts |
| CHT-002 | 100% stacked bar and column charts |
| CHT-003 | Combination charts |
| CHT-004 | Mekko / Marimekko charts |
| CHT-005 | Scatter and bubble charts |
| CHT-006 | Histogram and Pareto charts |
| CHT-007 | Small multiples |
| CHT-008 | Radar and heatmap visualizations |
| CHT-009 | Gantt charts and timelines |
| CHT-010 | Pie and doughnut charts |
| CHT-011 | Waterfall and bridge charts |
| CHT-012 | Line and area charts |
| CHT-013 | Stacked bar and column charts |

### Intelligence, labels and annotations (INT-001–INT-011)

| ID | Feature |
|---|---|
| INT-001 | Smart labels |
| INT-002 | Totals and subtotals |
| INT-003 | Difference arrows and comparisons |
| INT-004 | Legends |
| INT-005 | Error bars and uncertainty indicators |
| INT-006 | Custom annotations |
| INT-007 | Collision avoidance and label placement |
| INT-008 | CAGR calculations and indicators |
| INT-009 | Axes, scales and breaks |
| INT-010 | Number formats |
| INT-011 | Data-driven colors |

### Editing and interaction (EDT-001–EDT-013)

| ID | Feature |
|---|---|
| EDT-001 | Drag handles and interactive adjustments |
| EDT-002 | Snapping and alignment |
| EDT-003 | Multi-selection |
| EDT-004 | Keyboard shortcuts |
| EDT-005 | Grouping |
| EDT-006 | Chart conversion |
| EDT-007 | Resize and automatic reflow |
| EDT-008 | Distribution tools |
| EDT-009 | Context menus |
| EDT-010 | Undo and redo |
| EDT-011 | Theme integration |
| EDT-012 | PowerPoint selection synchronization |
| EDT-013 | Preservation of user overrides |

### Data integration (DAT-001–DAT-010)

| ID | Feature |
|---|---|
| DAT-001 | Embedded data editor |
| DAT-002 | Linked-data refresh |
| DAT-003 | Formulas and calculated values |
| DAT-004 | Database and API connectors |
| DAT-005 | Data provenance |
| DAT-006 | Data reconciliation |
| DAT-007 | Scheduled data updates |
| DAT-008 | CSV import |
| DAT-009 | Broken-link detection and recovery |
| DAT-010 | Excel range linking |

### Slide productivity (SLD-001–SLD-011)

| ID | Feature |
|---|---|
| SLD-001 | Agenda generation |
| SLD-002 | Harvey balls |
| SLD-003 | Process diagrams |
| SLD-004 | Templates |
| SLD-005 | Batch formatting |
| SLD-006 | Presentation quality assurance |
| SLD-007 | Slide library |
| SLD-008 | Corporate themes |
| SLD-009 | Icons |
| SLD-010 | Checkboxes |
| SLD-011 | Editable tables |

### Enterprise (ENT-001–ENT-010)

| ID | Feature |
|---|---|
| ENT-001 | Template-driven reports |
| ENT-002 | REST APIs |
| ENT-003 | Role-based access control |
| ENT-004 | Enterprise deployment |
| ENT-005 | Offline behavior |
| ENT-006 | Licensing |
| ENT-007 | Telemetry |
| ENT-008 | Single sign-on |
| ENT-009 | Audit trails |
| ENT-010 | Batch report generation |

## AI requirements (AI-001–AI-030)

| ID | Feature |
|---|---|
| AI-001 | Natural-language text-to-chart |
| AI-002 | Chart type recommendation |
| AI-003 | Conversational chart editing |
| AI-004 | Chart conversion assistant |
| AI-005 | Label optimization |
| AI-006 | Chart explanation |
| AI-007 | Natural-language data mapping |
| AI-008 | Data quality analysis |
| AI-009 | Governed natural-language data queries |
| AI-010 | Trend detection |
| AI-011 | Forecast assistance |
| AI-012 | Data reconciliation assistance |
| AI-013 | Text-to-slide generation |
| AI-014 | Slide layout recommendation |
| AI-015 | Corporate style application |
| AI-016 | Slide redesign |
| AI-017 | Narrative and storyline creation |
| AI-018 | Editable diagram generation |
| AI-019 | Executive summaries |
| AI-020 | Insight extraction |
| AI-021 | Presentation critique |
| AI-022 | Speaker notes |
| AI-023 | Audience adaptation |
| AI-024 | Source and claim verification |
| AI-025 | AI report builder |
| AI-026 | Multistep task planning |
| AI-027 | Natural-language automation |
| AI-028 | Change-impact analysis |
| AI-029 | Sensitive data protection |
| AI-030 | AI action audit and approval |

> **Traceability caution:** This index captures the planned feature families and identifiers. For precise functional behavior, error handling, acceptance tests and traceability, use the DOCX as the controlling specification. If this index differs from the DOCX, reconcile and update the index in a reviewed pull request.

## Milestone sequence

| Milestone | Focus | Exit gate |
|---|---|---|
| S0 | Windows VSTO feasibility, editable chart insertion, save/reopen | Validated in actual Windows PowerPoint |
| S1 | Canonical chart schema, geometry and persistence | Schema and deterministic unit tests |
| S2 | Core chart families | Editable chart acceptance tests |
| S3 | Labels, annotations and advanced chart behavior | Visual regression and interaction tests |
| S4 | Excel, CSV and enterprise data connectivity | Refresh, provenance and failure tests |
| S5 | Slide productivity and templates | Slide editing and layout tests |
| S6 | On-premises services and API | Security, authorization and contract tests |
| S7 | Azure OpenAI gateway | Policy and schema validation tests |
| S8 | AI-assisted workflows | Human approval and deterministic action tests |
| S9 | Enterprise controls and governance | Audit, identity and deployment tests |
| S10 | Performance, compatibility and production hardening | Production release criteria |

## Codex starting instructions

**Execution update — 21 October 2026:** The owner explicitly accepted S0 and authorized starting S1. This permits progression while the unexecuted S0 Windows acceptance checks remain unverified. The first S1 domain/contracts profile is under review; see [scope](../implementation/S1-domain-contracts.md) and [execution evidence](../evidence/S1-2026-10-21.md). It does not imply acceptance of any full chart/editing feature. The controlling DOCX section 13 places geometry, rendering and command execution in later slices; its sequence governs the S1 scope despite the broader shorthand in this index.

1. Read `AGENTS.md`, `docs/DECISIONS.md`, `docs/implementation/S0-feasibility.md`, and the full DOCX requirements.
2. Work only on **S0** until its acceptance tests pass in Windows PowerPoint.
3. Keep unknown Office versions, Azure tenant/model configuration, database choices and enterprise policies explicitly unresolved.
4. Do not assert completion based only on mocked or host-independent tests.
5. Open a reviewable pull request with changes, tests, evidence, limitations and outstanding decisions.
