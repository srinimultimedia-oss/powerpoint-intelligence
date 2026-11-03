# Checkmarx integrity follow-up — report 2 November 2026 13:42

## Scan and scope

The supplied full report (scan 13:38 Asia/Dubai, ID `544cab95-7e5b-4d76-b64c-a4ba3244db44`) has **zero active path-traversal findings**, **six detailed High PowerShell results**, and **one High supply-chain summary**: seven High total. Summary taxonomy is five Download of Code Without Integrity Check findings and one Inclusion of Functionality from Untrusted Control Sphere. All six detail records are grouped as CWE-829 without descriptions; both classifications and original IDs are retained here and in JSON. No result was excluded because its grouping is inconsistent.

Source changes implement controls for two cache-authenticity gaps and two signer/build-boundary findings. Two fingerprint-only flows have no download or executable sink and need scanner/reviewer validation as potential false positives; they are not suppressed or unilaterally marked resolved. A confirming scan is required for all six original flows. The supply-chain result has no original rule/ID/location in this report. The owner supplied `/scripts/Piidx-TestSuite.ps1` and `/scripts/Source-Fingerprint.ps1`; these match detailed SAST records but do not identify the separate 2ms rule. It remains unverified, not assumed to be a duplicate or a secret.

## Controls and threat boundary

The trusted inputs are the approved current source checkout, the normal Windows build-user environment and authenticated installed toolchain. Cache directories/receipts are not independent authorities. This does not authenticate an attacker-modified checkout or a compromised build user/Windows trust store.

- `Build-Integrity.ps1` hashes a single locked file stream; source fingerprint values retain their prior algorithm and ordering. Hashing code bytes does not execute them or prove their publisher.
- PIIDX now compares cached assembly bytes with a deterministic rebuild of current trusted source using a Microsoft-authenticated compiler, checks source stability and loads exactly the captured verified bytes. An already-loaded type requires matching verification state. Missing/unreceipted/changed cache gates remain; the missing PIIDX DLL was not recovered.
- Scene and template runners independently reproduce current-source artifacts in a verification-only build-user temp directory outside the cache. Those copies are never launched, installed or loaded. Cached code is compared and kept read-locked until execution finishes. Temp cleanup deletes only explicitly generated files, without recursive deletion. Existing executable paths remain stable. Cache recipe versioning and deterministic path mapping permit original/snapshot reproduction; source-review/quarantine gates remain enabled.
- The scene dependency is checked independently of its receipt: valid Microsoft Authenticode/catalog signature, Newtonsoft.Json identity/version/public-key token, and a locked reference throughout build/run. This matches the verified reference on this computer. A different machine with a differently signed/unsigned reference must supply approved independent dependency provenance; do not bypass the gate or infer publisher trust from a locally generated checksum.
- Compiler/MSBuild/vswhere require a valid Microsoft organization signature plus expected filename/OriginalFilename role. No broad trusted-publisher list or Windows trust policy was installed.
- easySetup uses an exact 40-hex thumbprint and typed read-only CurrentUser/My store search, not an arbitrary provider path. It requires validity dates, private key, explicit code-signing EKU and normal Windows chain/revocation success before building. Canonical thumbprints prevent MSBuild property injection. Existing final Valid signature checks and prior-installer preservation remain.

This relies on existing ACLs for trusted checkout, toolchain and private build-user temporary storage. File leases prevent ordinary cache replacement during use; no claim is made of complete handle-based isolation against a privileged namespace-race attacker. No dependency upgrade, Checkmarx suppression/custom sanitizer declaration, execution-policy bypass, trust-store or antivirus-policy change was introduced. Native PowerPoint/data/rendering source was not changed.

## Original finding register

| Original ID | Severity | Summary classification | Assessment | Source -> sink | Changes / validation |
|---|---|---|---|---|---|
| -440452681 | High | Inclusion / CWE-829 | Confirmed cache authenticity gap | `/scripts/Piidx-TestSuite.ps1:11` -> `/scripts/Piidx-TestSuite.ps1:26` | scripts/Piidx-TestSuite.ps1; scripts/Build-Integrity.ps1 |
| 1619924552 | High | Download / CWE-494 | Potential false positive: hashing, not code acquisition/execution | `/scripts/Source-Fingerprint.ps1:18` -> `/scripts/Source-Fingerprint.ps1:19` | scripts/Source-Fingerprint.ps1; scripts/Build-Integrity.ps1 |
| 1611841882 | High | Download / CWE-494 | Potential false positive: hashing, not code acquisition/execution | `/scripts/Source-Fingerprint.ps1:7` -> `/scripts/Source-Fingerprint.ps1:10` | scripts/Source-Fingerprint.ps1; scripts/Build-Integrity.ps1 |
| 1735884981 | High | Download / CWE-494 | Confirmed cache authenticity gap | `/scripts/Scene-RunnerCache.ps1:15` -> `/scripts/Test-SceneContracts.ps1:58` | scripts/Test-SceneContracts.ps1; scripts/Scene-RunnerCache.ps1; scripts/Build-Integrity.ps1 |
| -1447597795 | High | Download / CWE-494 | Conditional signer authority; not a download flow | `/easySetup/Build-EasySetup.ps1:14` -> `/easySetup/Build-EasySetup.ps1:78` | easySetup/Build-EasySetup.ps1; scripts/Build-Integrity.ps1 |
| 156779163 | High | Download / CWE-494 | Conditional build property injection / signer authority | `/easySetup/Build-EasySetup.ps1:14` -> `/easySetup/Build-EasySetup.ps1:27` | easySetup/Build-EasySetup.ps1; scripts/Build-Integrity.ps1 |

## Executed validation

Final source SHA-256: `6D4B966CB5F7A07544FE087842E325B14E06319FD208ED721CEC6B32D8737915`.

- Full local suite: **23 PASS / 1 FAIL**, `artifacts/validation-20261102-145526-7c15431e/results.json`, with matching `source-snapshot.json`. The failure remains the historical missing PIIDX DLL; no cache fallback/recovery was used. Thus full PIIDX loader runtime acceptance is not claimed.
- **20 build-integrity assertions**: cache plus forged receipt agreement still fails against independent rebuild; altered binary rejection; read-lock write refusal; verified bytes survive changed DLL path; constant-return fixture loads from checked bytes; malformed/semicolon/newline thumbprints rejected; signed Microsoft tool/role checks. Only an inert fixture is loaded, not a replacement PIIDX runner. Evidence: `artifacts/integrity-security-final.log` and aggregate security log.
- Exact valid-format but missing certificate lookup fails closed, and all modified entry points parse. No private key or certificate was created to pass tests.
- Scene checks on AnyCPU/x86/x64: **154 assertions + 73 cross-language round trips per architecture** and malicious corpus rejection, with fresh source proof and locked execution artifacts. Aggregate logs include the verification message. Current installed publisher/assembly identity for the dependency passed.
- Template checks: source proof enabled, 36 fidelity-policy assertions, 520 previews and 10,400 palette/layout combinations pass offline. Native Office insertion/edit/save/reopen is not inferred from these results.
- `easySetup/Test-EasySetup.ps1 -SourceOnly`: **8 setup policy + 3 resource-reader assertions pass** (`artifacts/integrity-setup-source.log`). No installer was published/installed; trusted private-key signing and deployment remain desktop gates. Existing certificate expiry/trust behavior was not bypassed.
- Existing corpus/path/diagnostic/Python security tests remain passing. No broader repeat testing was performed after these checks without a new concern.

## Remaining acceptance and regression considerations

Run a new Checkmarx scan and compare the six original records, new records and actual flow classification. Read-only fingerprinting results require reviewer evidence rather than a blanket suppression. Obtain the missing supply-chain rule/ID/location before remediation of that summary can be claimed.

The PIIDX cached assembly remains missing with a prior receipt. Resolve it through the approved scanner/source-review recovery process; then execute the full PIIDX suite to validate in-memory loading and canonical deterministic builds. Do not delete the receipt or restore/rename a quarantined binary.

From a normal desktop token, verify approved private-key build/signing with the stricter certificate gates, run the full interactive suite and verify native PowerPoint behavior. Source changes make the old installer stale; publish only through the verified normal signing workflow. Certificate revocation/trust/network failures and alternate dependency publishers now fail closed with an actionable message. Normal chain policy may require the organization's ordinary validation connectivity. These are deliberate trust boundaries, not silently ignored failures.
