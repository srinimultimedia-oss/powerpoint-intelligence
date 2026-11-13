# Architecture diagram index

The architecture diagrams are maintained with their interpretation and source evidence:

| View | Document | Meaning |
|---|---|---|
| C4 context | [Enterprise](enterprise-architecture.md) | Author, PowerPoint, local sources/catalog and provider-preparation boundary |
| C4 containers/processes | [Solution](solution-architecture.md) | Office/VSTO process, local storage, development Node/compiler processes |
| C4 components | [Components](component-architecture.md) | Lifecycle/UI, feature handlers, codecs, PIIDX and editable surface |
| Deployment | [Deployment](deployment-architecture.md) | Dual-bitness signed payload build and per-user VSTO activation |
| Execution sequences | [Sequences](../10-workflows/sequence-diagrams.md) | Host events, linked writes, library operations and teardown |
| Activities/trust boundaries | [Flows](../10-workflows/activity-diagrams.md) | Validation, current-window mutation and failure branches |
| Data/ERD | [Data classes](../08-data-design/class-diagrams.md) | Actual model and PIIDX entities |

All graphs describe inspected code boundaries, not measured throughput, HA topology, unimplemented Azure invocation or third-party feature parity. Mermaid fences are inventoried and structurally checked. Full syntax parsing/rendering is separately reported; lack of a local Mermaid runtime must not be called a parser pass.
