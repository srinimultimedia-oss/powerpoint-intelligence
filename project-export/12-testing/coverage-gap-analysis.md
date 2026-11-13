# Test coverage gaps

There is no collected line/branch coverage denominator in the repository. Source/test association is traced in test-feature-traceability.csv, but an existing related test path is not proof every command branch or API was exercised. The current constrained run intentionally covers deterministic source/metadata behavior only.

| Gap | Evidence and risk | Required validation |
|---|---|---|
| Office estate | Profiles share API16; product/build/architecture matrix unproven by configuration tests | Interactive2016/2019/365 builds andx86/x64 package loading, not just compiler targets |
| Native renderer breadth | Generic scene DTOs/TS support more geometries than native clustered rectangle adapter | Per-family capability tests and deliberate rejection before mutation |
| V3 production | V3 source omitted compile lists; no direct codec harness found | Explicit build/migration decision and direct boundary tests before wiring |
| Full feature command behavior | Catalog/ribbon shape checks exceed native acceptance coverage | Every routed command, unsupported selections, undo, cancellation and persistence |
| Gallery/native templates | Procedural layouts and original source imports are different modes;16-original historical sample | Full intended native corpus/formatting edit/save/reopen and license-aware asset tests |
| DPI/foreground/locale | Context controls depend on native focus, selection and WinForms layout | Multi-monitor scaling, display transitions, keyboard/accessibility and locale matrix |
| Data link refresh | Primary-axis uniform category charts only in mapped refresh | Moved/header/shape changes, XY/pie/secondary axes, long-lived Excel lifetime if expanded |
| AI/enterprise | No implementation, so labels/schema cannot be end-to-end tests | Authenticated transport/governance/approval/audit tests only after implementation |
| Build supply chain | TestS1 fallback download lacks independent digest; PIIDX scanner evidence filter mismatch | Tampered package tests and positive detection fixtures covering actual runner paths |
| Host-globalization fingerprint | NLS/ICU ordering can yielddifferentpackaging fingerprints | Cross-PS5/PS7 ordinal canonical vectors; no recipechange made here |

This export's completeness metric measures documentation/traceability and bounded validation, not product acceptance. Planned/missing modules remain visible rather than being promoted to verified implementations.

## Semantic review refinements

The complete64 authored tests/ source files and7 package tests are mapped; full-body reviewer CSVs distinguish helpers from runners. DemoContractTests checks scenario text presence, not native command acceptance. Most native suites call source implementations compiled into the test executable; installed-marker/context checks have narrower explicit scopes. ElementsOnly also runs context/font suites. NativeFormattingTests has a cancel-path `touched` variable that is never changed, making that portion of its non-mutation oracle vacuous. MappedExcelOfficeTests does not explicitly Quit its separately created Excel instance, and DataLinkManagerTests retains OS-temp fixtures. Some broad reopen checks assert only counts; focused tests inspect deeper metadata. These are documented test gaps, not changed code or newly executed cases.
