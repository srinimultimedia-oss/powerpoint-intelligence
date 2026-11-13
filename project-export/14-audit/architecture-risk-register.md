# Architecture risk register

| Risk / assessed priority | Actual boundary | Impact and mitigation |
|---|---|---|
| AR-01 High — Office UI/COM coupling | Most feature logic mutates native objects on interactive UI STA | Long work/Excel lifetime failures affect user workflow; keep bounded parsing/native commit verification, isolate pure work and test owned rollback |
| AR-02 High — acceptance scope drift | Product catalog/requirements exceed native verification | Maintain source/runtime/estate matrices; never infer compatibility from API16 or DTO type count |
| AR-03 High — inherited installation trust | Local development signer versus estate publisher/VSTO load | Approved distribution publisher lifecycle, target-user trust and rollback tests; no policy bypass |
| AR-04 High — external-service architectural absence | UI/profile/schema only; auth/network/data policy unresolved | Implement service security and transport only after owner decisions, with contract and privacy controls |
| AR-05 Medium — editable object/canonical divergence | Native manual edits may disagree with tags/scene/canonical metadata | Preserve native edits and report explicit divergence; do not silently overwrite canonical or user formatting |
| AR-06 Medium — local index confidentiality/authenticity | PIIDX contains text/paths and unkeyed integrity digests | Define permissions/retention and threat model; digest is accidental-tamper detection, not adversarial authenticity |
| AR-07 Medium — generated artifacts and tooling trust | Multiple cache/recovery and package paths | Source-reproduction, signature/digest checks, scope-complete scanner evidence and immutable receipts |
| AR-08 Medium — synchronous catalog/package work | Local metadata extraction and native previews scale with files/slides | Bound aggregate work, cancellation and queue limits; measure real workloads, not unit runtime |
| AR-09 Medium — source/procedural/third-party content modes | Different licenses/provenance/native behavior | Preserve source references/license constraints; do not ship proprietary assets in documentation ZIP |
| AR-10 Medium — platform dependencies | Office PIAs/Framework4.8, winsqlite3/FTS5 and filesystem policy | Explicit preflight/capability rejection and target-estate testing before deployment claims |

Evidence: `apps/powerpoint-vsto/ThisAddIn.cs:37`; `apps/powerpoint-vsto/DataSources.cs:166`; `src/Chart.Renderers/PowerPointSceneSurface.cs:147`; `src/Library.Core/NativeSqlite.cs:1`; `src/Library.Core/PiidxSchema.cs:1`; `easySetup/Build-EasySetup.ps1:1`; `docs/DECISIONS.md:1`. Risk ratings are this documentation review's prioritization, not vulnerability scan IDs or quantified likelihood.
