# OPS-08 — Offline alignment repair utility

**Evidence category:** source verified; execution status is stated separately. Baseline includes the actual inspected checkout, not an assumption that every working byte equals HEAD.

## Design

| Dimension | Implementation-grounded description |
|---|---|
| 1. Purpose | Generate a new centered-text PPTX copy with bounded ZIP/XML work and preservation verification; never edit original/open deck. |
| 2. Implementation status | Implemented Python/lxml offline tool; native rendering acceptance explicitly false in sidecar. |
| 3. Internal components/classes/interfaces | Repair-TemplateAlignment bounds/contains/fix_slide/repair; RepairTemplateSafety Limits/Budget/preflight_zip/inspect_archive/read_member/parse_xml/validate_repair/publish_new. |
| 4. Important functions/methods | CLI --input/--output; repair(source,output,limits); fix_slide returns XML/count/fittedlabels; explicit require remains active under python -O. |
| 5. Inputs and outputs | Selected PPTX ZIP and new output path; outputs new deck plus .validation.json with counts/preservation and powerPointRenderingVerified=false. |
| 6. Dependencies | Python pathlib/zipfile/io/crypto/tempfile/lxml, no PowerPoint COM. |
| 7. Internal workflow | Validate new targets/source distinction; preflight central directory/limits/hash; stage package; center body anchor and paragraph alignment; fit single labels within eligible containers; verify all member names/text/untouched parts/timing/transitions; rehash unchanged source; no-overwrite publish deck/sidecar. |
| 8. External integrations | Filesystem only; XML network/entity loading disabled. No Office/rendering call. |
| 9. Exceptions and error handling | Any size/node/depth/work/time/name/duplicate/encryption/DTD/preservation/mutation/publication violation rejects; partial own output rolled back and stagefiles removed. |
| 10. Configuration | Limits defaults:512MBinput/archive,10000 members,64 MBmember/directory,40 MBXML,1 Mnodes,depth128,4 GBprocessed,25 Mwork,120 seconds. |
| 11. Security | Reads actual bounded member bytes, rejects unsafe/duplicate package names before transformation; external entities disabled/no network; new destinations only. |
| 12. Tests | 22 test_repair_template_alignment cases cover preservation, no-overwrite, budgets/compressedbomb, mutation, hostile XML/names, rollback and -O safety. |
| 13. Known limitations | Geometry fit is heuristic for single labels/simple shapes; XML-preserved transitions do not prove rendering/playback; no current native visual QA claimed. |
| 14. Technical debt | No full Office font/layout metrics; very complex/transformed/grouped text may need manual validation after offline repair. |
| 15. Extension points | Expand bounded geometry/layout test corpus and native visual comparator separately; keep preservation and no-overwrite guarantees. |

## 16. Complete owned source map

| Source | Analysis | Responsibility/declarations |
|---|---|---|
| `scripts/Repair-TemplateAlignment.py:1` | Complete text read; no runtime execution | Script/project entry; its actual flow and responsibilities are documented above. |
| `scripts/RepairTemplateSafety.py:1` | Complete text read; no runtime execution | Script/project entry; its actual flow and responsibilities are documented above. |
| `tests/security/test_repair_template_alignment.py:1` | Complete text read; no runtime execution | Script/project entry; its actual flow and responsibilities are documented above. |

## 17. Handover and maintenance checks

Review the exact source/hash/line counts in `09-security/reviewed-files.csv` and this directory’s `source-module-mapping.csv`. Do not substitute historical PASS logs for current execution. Preserve unrelated module behavior and tests before modifying a boundary; new methods must have typed inputs, explicit error semantics and a corresponding negative test. See the dedicated AI/API/security/deployment documents for this module’s external contract.
