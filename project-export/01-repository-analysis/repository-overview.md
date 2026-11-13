# Repository overview

**VERIFIED:** baseline commit `ffaec352cbdca666f69f3c402d334108e4115933` (3-Nov-2026), inspected live checkout `C:/Srini/powerpoint-intelligence/powerpoint-intelligence` on13-Nov-2026. The inventory excludes project-export to avoid recursively documenting itself. Every discovered regular file is listed; links are classified without traversal. Git/compiled/vendor/artifact data is counted but not confused with authored implementations.

- Repository inventory: **30,301 file/link records**, **2,736,823,893 bytes**.
- Authored code: **204** `.cs/.ps1/.py/.ts/.mjs/.tsx` files outside generated/vendor/VCS/export areas.
- Relevant decoded source/config/schema/documentation text: **339** files.
- Symbol locator: **1,833 lexical candidates**, not an AST-derived API count.
- Inventory read/hash errors: **0**.

| Area | Actual responsibility | Evidence interpretation |
|---|---|---|
| apps/powerpoint-vsto | VSTO Framework4.8 add-in, WinForms UI, COM features, codecs, local gallery | Active project compile items determine production inclusion |
| src/Chart.Domain | v2 canonical JSON, S0 import; authored V3 files | V3 excluded from active projects |
| src/Chart.Renderers and Shared.Contracts | strict scene contracts/hash, clustered native feasibility renderer, bounded JSON | Native capabilities narrower than generic scene DTO types |
| src/Library.Core | independent PIIDX SQLite/catalog, source extraction, preview/config/watch | Local searchable catalog, not proprietary index loading |
| packages/chart-core/domain/engine | deterministicJS/TS, schema validation and offline scene/SVG tools | Not an OfficeJS frontend or deployed network worker |
| services/ai-gateway | JSON schema and explicit unimplemented README | Planned service only |
| scripts and easySetup | compilation/test/cache/security/diagnostics and dual-bitness signed setup | Some paths write artifacts/install state and are not run during this export |
| tests/.github | custom Console, Node and unittest suites and limited CI | Test existence is not execution/line coverage |
| docs | requirements, decisions, prior evidence and remediation records | Historical/design records are secondary to source |
| artifacts | cached runners, receipts, snapshots, prior evidence/build products | Generated and historical; not additional source modules |

`git-baseline.json` distinguishes235 byte differences from HEAD:234 are CRLF-only; the sole remaining tracked difference is generated `easySetup/README.md`. Source is therefore pinned by actual per-file SHA256 as well as commit, without implying functional uncommitted changes. Untracked ignored deployment assets/dependencies remain in the inventory. No existing export was present, so no prior export was overwritten.
