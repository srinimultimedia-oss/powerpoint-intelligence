# Activity diagrams

## Validation before document mutation

```mermaid
flowchart TD
  A[User action] --> B{Host available and compatible}
  B -- No --> X[Report disabled or unavailable]
  B -- Yes --> C{Mutation lease acquired}
  C -- No --> Y[Return without competing write]
  C -- Yes --> D[Resolve current presentation and selection]
  D --> E[Read and validate canonical data or input table]
  E --> F{Supported and unambiguous}
  F -- No --> Z[Preserve native objects and report]
  F -- Yes --> G[Start native undo boundary where implemented]
  G --> H[Perform scoped native edits]
  H --> I[Verify commit and persist association]
  I --> J[Update status and context observation]
```

`apps/powerpoint-vsto/ThisAddIn.cs:133`; `apps/powerpoint-vsto/PowerPointHost.cs:48`. Undo grouping and transaction rollback differ by module; the graph is a shared validation pattern, not a claim every handler implements the same atomicity.

## Context properties and stale target prevention

```mermaid
flowchart TD
  A[Selection or timer request] --> B[Capture on Office UI STA]
  B --> C{PowerPoint owner foreground}
  C -- No --> H[Hide property box]
  C -- Yes --> D[Classify applicable properties and target identity]
  D --> E[Publish immutable observation snapshot]
  E --> F[User chooses property edit]
  F --> G{Selection identity still matches}
  G -- No --> R[Reject stale mutation]
  G -- Yes --> M[Apply only supported properties]
```

`apps/powerpoint-vsto/ContextPropertiesController.cs:1`; `apps/powerpoint-vsto/ContextSelectionBridge.cs:1`; `apps/powerpoint-vsto/ContextActionRouter.cs:1`. Foreground focus is a Windows prerequisite, not a reason to change focus/security policy. Table cell text, group text, chart properties and indicators have different applicability.

## Local catalog refresh

```mermaid
flowchart TD
  A[Configured enabled source root] --> B[Enumerate supported packages within policy]
  B --> C{Fingerprint changed or new}
  C -- No --> D[Retain current indexed presentation]
  C -- Yes --> E[Parse bounded presentation metadata]
  E --> F[Begin catalog transaction]
  F --> G[Replace presentation slides objects tags and FTS rows]
  G --> H[Commit or roll back error]
  H --> I[Report refresh progress and missing files]
```

`src/Library.Core/PiidxCatalog.cs:1`; `src/Library.Core/PresentationSource.cs:1`. Watch notifications/debounce request refresh; they are not a guarantee of enterprise scheduled updates or real-time external synchronization.
