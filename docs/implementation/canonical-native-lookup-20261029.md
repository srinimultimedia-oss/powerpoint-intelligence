# Canonical native metadata lookup — 29 October 2026

Desktop input: deployment 1.0.0.6, piidx4 marker; artifacts/demo-features-252195dedd1245af833647e6f25a002b.

Confirmed native progress: 30 chapter-agenda assertions, 34 workflow assertions including edit/save/reopen/PDF, and the previously failing outline-color assertion pass. Font and productivity suites pass. Remaining failures are preserved as native acceptance gates.

## Changes

Text/indicator readers still filtered by the native CustomXMLPart.NamespaceURI property. They now use a central bounded selector that normalizes and deduplicates GUID identities, reacquires each part through its owner's CustomerData.Item(id), then compares the serialized root namespace. XML update verification, grid text lookup, source references and connector lookup use the same selector. Distinct canonical part IDs remain ambiguous even when their XML is identical; unknown or malformed ownership is never silently removed or adopted.

Microsoft documents CustomerData.Item as GUID lookup and documents copying customer data to a new CustomXMLPart for the copied object: https://learn.microsoft.com/en-us/office/vba/api/powerpoint.customerdata.item and https://learn.microsoft.com/en-us/office/vba/api/powerpoint.customerdata . The change avoids relying on an enumerator's cached namespace state; native confirmation is still required.

New text elements and duplicates receive unique names containing native shape IDs. The multi-selection test resolves current native shape indices from IDs and activates the correct slide, rather than selecting by potentially duplicated names.

The connector no-op test now saves its owned fixture to disk and verifies a clean Saved baseline before testing reflow. Earlier traces already showed Saved=false before reflow entered and zero planned changes. No Saved flag is forced true to conceal a mutation.

Linked-chart commit polling reads native state before a redraw and refreshes at most once, rather than refreshing repeatedly throughout an eight-second polling loop. If it still fails, the error reports expected/actual counts and mismatch indices/types, never source values. Commit failures and RPC failures still fail acceptance. Chart insertion cleanup retains its original error; nested disconnections are recognized so later commands remain BLOCKED rather than being attempted against a dead server.

Marker: PPI-DEMO-RIBBON-20261029-v17-piidx5-canonical-parts.

## Verification

Full non-interactive suite: artifacts/validation-20261029-165540-7cfbc429/results.json, 23/23 PASS. The final extraction of canonical selection into a pure helper was checked separately with six ownership/schema tests, bringing XML/locale/coordination contracts to 44 assertions. These cover duplicate GUID spelling, distinct same-schema ambiguity, foreign namespaces, invalid IDs, lost membership and bounded enumeration. Final ribbon contracts and demo-test compilation pass; all three production architecture compilation logs are in artifacts/native-lookup-20261029.

These are not new native PowerPoint passes. Desktop installation and Test-DemoFeatures.ps1 -InteractiveOffice remain required. If actual distinct canonical parts remain duplicated, the code deliberately preserves them and reports ambiguity instead of choosing one arbitrarily. The content-free chart diagnostics are intended to identify the remaining commit mismatch if polling still fails.