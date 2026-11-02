# Full desktop validation follow-up — 29 October 2026

Evidence: artifacts/validation-20261029-230052-bd2ddae8, deployment 1.0.0.10. The prior focused CSV diagnostic passed; it did not establish mapped Excel refresh or the complete native suite.

- Template fidelity never reached insertion: the stable runner resolves templates relative to its assembly and had no explicit asset configuration. Snapshot assets exist. Test-TemplateLibrary.ps1 now creates a runner config pointing explicitly to the current source snapshot template folder. It reuses the same executable and does not change its binary receipt. Compile-only preparation verified this configuration path; rendered fidelity remains untested in this session.
- Mapped Excel refresh failed with zero chart points before workbook close. Unlike the successful CSV test, source readers and the mapped fixture writer activate Excel and explicitly quit that application after closing their workbook. This process-wide shutdown is unsafe when ownership of an Excel automation server used by chart embedding is not established. Cleanup now closes the owned source workbook and releases its acquired application COM reference instead of calling Quit. This change is a candidate remediation, not proof that those calls caused the observed Excel/PowerPoint crashes. Native lifecycle/resource behavior remains to be validated.
- S0 office-1 reached chart update and failed with RPC unavailable; office-2 refused to run because a presentation remained open. No third run was executed. Do not count refusal, blocked commands or historical successes as new passes. Crash records include Excel and PowerPoint chart.dll errors; they do not establish the offending caller.
- The diagnostic observer is now an explicitly settable thread-static-backed property, removing the production CS0649 warning without suppressing warnings.

New focused regression: scripts/Test-DemoFeatures.ps1 -InteractiveOffice -MappedExcelOnly. It runs the real explicit non-A1 mapping, changed source refresh, repeated refresh association, changed-header refusal, and save/reopen assertions without the full feature catalog. It writes mapped-excel-result.txt and returns failure on any exception. It uses the existing stable DemoOfficeTests.exe.

Test-DemoFeatures.ps1 compiled successfully and Test-TemplateLibrary.ps1 -CompileOnly reused the stable runner. Desktop PowerPoint tests cannot be executed by this AppContainer session. No security settings or exclusions were changed; no user presentations were closed by the agent. Full acceptance remains open.
`nLocal verification after the changes: artifacts/validation-20261029-230930-4d4cddeb, all 23 noninteractive checks PASS, including three architecture builds. This does not establish native Office acceptance.

## Mapped source registry reopen correction

Desktop evidence demo-features-38f1aa13b75c47ebaf6867fd805215e6 passes the first four mapped Excel assertions: non-A1 mapping, changed-source refresh, repeated association reuse and changed-header refusal. It then fails DataSources.Read after save/reopen. Independent inspection of mapped-range.pptx confirms one source identity in two identical XML payloads. The deck registry and the shape-attached source reference are both exposed through presentation.CustomXMLParts after reopening.

DataSources.Read now gathers shape CustomerData ownership (including grouped children, with traversal bounds) and excludes those parts from deck registry enumeration. Genuine duplicate deck entries are still rejected. Three pure regression assertions cover attached-reference exclusion, duplicate-deck rejection and refusal to treat orphan shape references as registry entries. No saved parts are deleted or rewritten by the reader.

The same desktop report establishes all 16 original template fidelity slides passing. This is coverage of those 16 slides, not all 497 assets. Mapped Excel save/reopen and full native acceptance remain pending execution of Test-DemoFeatures.ps1 -InteractiveOffice -MappedExcelOnly against the corrected source. No reinstall is necessary to run the standalone test.
