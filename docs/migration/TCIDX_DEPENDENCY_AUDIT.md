# TCIDX dependency audit — baseline

Baseline: `78b9f87ef2218d6d19d7e8049d0f223281e9ed76`. The merged implementation is mostly untracked; inventory CSV includes those files separately. No production modification preceded this report.

## Confirmed dependencies and replacements

| Location | Class | Verified behavior | Required replacement / verification |
| --- | --- | --- | --- |
| apps/powerpoint-vsto/TemplateLibrary.cs:475 | A | Optional fallback constructs the installed vendor templates directory. The explicit configured local directory currently takes priority. | Remove installed-vendor discovery after independent library registration is integrated; test absent vendor installation. |
| apps/powerpoint-vsto/TemplateLibrary.cs:334–455 | A | LocalSlideIndex directly extracts Open XML from PPTX/POTX. LocalLibraryCache serializes a version-1 XML catalog. No TCIDX reader found. | Reuse supported extraction behavior while adding stable native slide IDs, objects, fingerprints and SQLite PIIDX persistence. Preserve native source insertion. |
| apps/powerpoint-vsto/LocalLibraryGallery.cs:48,72,130–139 | D | Single remembered XML cache, in-memory filtering and full reindex; native previews run from the UI timer. | Connect registration, SQLite/FTS5 search, incremental refresh, cancellation and error reporting. Keep COM previews on UI STA; persist validated preview references. |
| apps/powerpoint-vsto/App.config | C | Explicit local source template folder. No TCIDX setting. | Preserve this as source-library migration input; add documented independent index/enterprise settings. |
| apps/powerpoint-vsto/PowerPointIntelligence.csproj | B | .NET Framework 4.8; Microsoft Office/VSTO and Newtonsoft references. No vendor DLL/COM reference found. | Add independent engine sources and tested SQLite provider; verify deployment payload. |
| scripts/Test-TemplateLibrary.ps1:36 | F | Native fidelity test chooses installed vendor templates when present. | Make source-library fixture paths explicit and independent; retain owner-approved fidelity policy. |
| apps/powerpoint-vsto/FeatureInventory.cs:78,95,105 | G | Descriptive feature inventory mentions the vendor; no library service invocation. | Preserve unrelated behavior; text alone is not an indexing dependency. |
| docs/implementation/*reference*, thinkcell-*; docs/requirements/* | E | Historical observations and requirements. | Retain with classification, never package research fixtures as runtime inputs. |
| artifacts/merge-*/previous, template-copy*.json | E | Historical provenance/backups. | Preserve; do not rewrite provenance or count these as active dependencies. |

The companion `dependency-findings.csv` contains every keyword/loading-reference location. Inventory records project/config/script content scans; binary assets were hashed and ZIP members inventoried, not claimed as manually reviewed. Framework COM activation and native process references require preservation, not blanket replacement. The deployed package has not been built or activated for this migration.

## Risks and gates

- Existing accepted native template insertion/fidelity and contextual formatting are regression gates. Source package copies are inputs, not a proprietary indexing dependency.
- SQLite must actually execute FTS5, transactions and integrity checks on each supported process architecture. OS-provider availability is an explicit compatibility gate, not inferred from Office version.
- Keep prior valid indexes and rollback metadata on cancellation, corruption or source changes. Never decode TCIDX for production migration.
- Existing AppContainer Office activation denial is a separate baseline; owner desktop activation previously succeeded. This migration does not claim to fix it.
- Scene x64 quarantine hold and historical runner receipts remain unchanged. No scanner exclusion, alternate runner recreation or security-policy changes are authorized here.

Status: discovery complete for enumerated baseline; replacement implementation and final coverage/runtime verification pending.