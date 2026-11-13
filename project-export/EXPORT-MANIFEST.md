# Export manifest

**Application:** NexWave / PowerPoint Intelligence. **Repository:** `C:/Srini/powerpoint-intelligence/powerpoint-intelligence`. **Baseline commit:** `ffaec352cbdca666f69f3c402d334108e4115933`. **Export version:** 1.0. **Generated:** 13 November 2026, Asia/Dubai; exact UTC times are in the validation records.

## Contents and baseline

All **81 explicitly requested deliverable files** are supplied, with **46 individual module designs**, API/native/event appendices, detailed feature designs and supporting CSV/JSON validation data. The package contains **126 Markdown documents**. [Documentation index](DOCUMENTATION-INDEX.md) lists them; [file inventory](14-audit/export-file-inventory.csv) records the portable files and checksums.

Repository discovery covers 30,301 records: 30,300 regular files and one historical test junction that was not followed. The authored-code denominator is 204 files; every file has a source-specific review record. Forty-six VSTO code files include targeted-body review with complete-text indexing rather than formal line-by-line proof. There are 339 relevant text files in total; supporting historical documents have explicitly different inspection levels. Source hashes, line counts and exclusions are preserved in the [repository inventory](01-repository-analysis/complete-file-inventory.csv) and [review coverage](01-repository-analysis/review-coverage.csv).

The current source baseline matches HEAD after Git line-ending normalization, except generated easySetup/README.md. The full before/after comparison confirms no regular files outside project-export were added, removed or changed. This export did not install software, upgrade dependencies or change application code, configuration, existing binaries or security settings.

## Coverage and validation

Coverage is explicitly scoped: 204/204 authored code files accounted for, 46/46 module designs, 190/190 feature entries (172 UI routes and 18 capabilities), 12/12 declared interface groups (one schema-only/planned boundary), and 71/71 test-source files mapped. Related tests are regression candidates, not instrumented feature coverage. [Completeness report](15-deliverables/DOCUMENTATION-COMPLETENESS-REPORT.md) explains all denominators and exclusions.

Current safe checks passed 229 assertions, TypeScript noEmit and syntax parsing of 50 PowerShell files. No current Office integration, signed installation, vulnerability scan or external-service execution is claimed. Source references, internal document links, anchors and Mermaid fence/type structure pass validation. A full Mermaid parser was unavailable; zero diagrams are claimed syntax/render verified. [Validation report](14-audit/documentation-validation.md) provides exact counts and limitations.

OpenAPI is deliberately inapplicable because no HTTP server or routes are implemented; the [OpenAPI note](07-api-design/openapi/README.md) explains why no fictional specification was created. AI profiles/schema remain scaffolds, V3 contracts are excluded from production compile lists, and source-backed security/test gaps remain recommendations rather than completed remediation.

## Portable delivery and integrity

The documentation archive is `15-deliverables/PowerPoint-Intelligence-NexWave-Technical-Design.zip`. It contains documentation and validation data only, with no application source files, dependencies, private configuration, credentials, binaries or licensed template assets. Local `_analysis`, `.unlazy`, authoring helpers and temporary test executables are excluded. Source citations remain repository-relative paths and lines; they are not copied source files.

[DOCUMENT-CHECKSUMS.sha256](15-deliverables/DOCUMENT-CHECKSUMS.sha256) hashes the selected documentation/data and file inventory. It excludes itself and external archive/hash receipts to avoid cycles. The finished ZIP checksum is in `15-deliverables/PowerPoint-Intelligence-NexWave-Technical-Design.zip.sha256`; CRC and entry comparisons are recorded in `14-audit/archive-validation.json`. Those two external delivery receipts and the ZIP itself are not recursively included in the archive. Document-to-document hyperlinks target packaged content.

**Status:** requested documentation delivered with explicit limitations. This is neither production acceptance nor scanner clearance. No prior export existed; future regeneration should preserve this baseline or use a versioned export folder.
