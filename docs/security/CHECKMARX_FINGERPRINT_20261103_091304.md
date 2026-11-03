# Checkmarx fingerprint flow — 3 November 2026, report 09:13

## Report facts

Scan `927ce31c-f55d-4fd8-ba2a-3a4327d8549f`, main, 09:07 Asia/Dubai; report created 09:13. Total: **two High** — one detailed SAST record plus one 2ms supply-chain summary. No active traversal or untrusted-functionality category appears. All six original PowerShell ID occurrences from the preceding report appear in its resolved tables (five CWE-494 and one CWE-829); this was compared programmatically, not inferred from the total.

Remaining SAST record **1072025071**, High, To Verify, CWE-494, Download of Code Without Integrity Check:

- Reported source: `scripts/Source-Fingerprint.ps1:19`, `Get-PpiAddinFingerprint`, extension metadata from Get-ChildItem.
- Reported destination: same file line 20, the Get-PpiContentHash call.
- Original result URL and exact source/sink content are in the JSON companion/report. The rule supplies no description and no query path beyond its name.

## Actual flow and assessment

1. Get-PpiAddinFingerprint selects filesystem entries, filters/sorts them, and obtains their relative filenames.
2. Get-PpiContentHash in Build-Integrity.ps1 opens each selected path with FileMode.Open, FileAccess.Read and FileShare.Read. It calls SHA256.ComputeHash on that stream, returns a hex digest, and disposes the stream/hash. It does not write, parse, import, compile, load or launch that selected file.
3. The caller aggregates relative-name/content-hash records and returns another SHA-256 digest. No path from this enumeration is passed to a network download or execution API in the reported chain.
4. Install-Local.ps1 uses this digest before/after its normal build to detect source edits. easySetup/Common.ps1 incorporates it into freshness accounting. These fingerprints are mutation/freshness evidence, not a claim to authenticate the authorship of an arbitrary checkout.

**Assessment: potential false positive / query-model misclassification for this reported flow.** Directory enumeration is a taint source, not by itself a CWE-494 defect. The documented destination is an integrity-hash generator, not a code-download/execution sink. A separately reported code-loading or remote-retrieval path would need its own analysis. The approved source checkout/build user remain explicit trust assumptions; this document does not claim that hashing alone authenticates downloaded source.

No production rewrite, command renaming, custom sanitizer declaration, scanner exclusion, suppression or Not Exploitable state change was performed. A Checkmarx reviewer must assess the complete query path before this record is marked resolved/non-exploitable. It remains **open / To Verify** in our register. Renaming the enumeration or hiding it behind a different API merely to remove a result would not constitute a verified security fix.

## Executed evidence

`scripts/Test-SourceFingerprint.ps1` exercises the actual production fingerprint functions against a known local inventory including an inert script body that would create a marker and throw if executed. **Nine assertions pass**:

- Digest matches the independently selected known file records/content hashes.
- Selected script content never executes; the marker remains absent.
- Only a SHA-256 value is returned and repeated input is deterministic.
- Generated bin output is excluded.
- General fingerprinting also treats the script as bytes.
- Changing source content changes both relevant fingerprints, without execution.

Evidence: `artifacts/source-fingerprint-flow-20261103.log`. No HTTP request or executable/assembly fixture is created by this test. Test source files are inert data, not invoked. Application, hashing algorithm and integrity/receipt/signing gates are unchanged.

The existing security suite also passes: 38 corpus assertions, 30 path-authority assertions, 31 diagnostic assertions, 20 build-integrity assertions and 22 Python tests (`artifacts/source-fingerprint-security-20261103.log`). Junction creation remains skipped in this restricted session. Native Office/private-key publishing and full cached PIIDX tests were not rerun or claimed passed; their previous acceptance limits remain.

## Checkmarx review handoff

For result 1072025071, inspect the query's complete source-to-sink graph and confirm whether the read-only custom hashing function is being treated as a download/execution sink. Review FileAccess.Read, SHA256.ComputeHash, return type and downstream freshness comparisons. If the full graph adds a different sink, retain the record and provide that path for remediation. Otherwise record a reviewer-supported false-positive decision under the organization's normal process. This handoff requests analysis, not automatic suppression or a blanket path/query exception.

The second High result is still a 2ms summary without its rule, original ID, file/line or secret type. The previously supplied script paths overlap SAST records and do not identify it. Do not assume it has been fixed or is a duplicate. Full Checkmarx compliance remains unverified until the remaining records have supporting evidence and scanner/reviewer acceptance.
