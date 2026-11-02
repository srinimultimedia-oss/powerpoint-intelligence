# easySetup resource-validation correction — 2 November 2026

The desktop reported an embedded-payload mismatch for published version 1.0.0.14. Read-only inspection in a fresh process established:

- EXE SHA-256 matches build.json: `ED84FC9F968A3DD6E11F330258CE4F11A378D4865D26C554F4418E1515178BED`.
- Embedded ZIP SHA-256 matches build.json: `9559CBD5039D33D74657B923B1CEEC93CCF26E8FECEA3A1B1626AEBDA930F909`.
- The resource stream begins at offset zero. There was no evidence of ZIP corruption in this published package.

The validator used Assembly.LoadFile on a fixed EXE path. An inert reproduction compiled two same-identity DLLs with different resource data, loaded the first from a stable path, atomically replaced that file and read it again. LoadFile returned the first resource while Assembly.Load(current bytes) returned the replacement resource. Thus a long-lived PowerShell session can compare an older loaded resource with a newer EXE hash/receipt.

Common.ps1 now hashes a resource from fresh immutable assembly bytes. Test-EasySetup.ps1 takes the same build mutex as publication, rejects an active build, holds a read-only non-write/non-delete-sharing file handle, hashes one bounded byte snapshot, validates the Authenticode signer/trust, and then hashes that snapshot's payload. Existing source freshness, EXE hash, payload hash and signer checks remain mandatory. No receipt or expected hash was edited to make a check pass.

Test-ResourceReader.ps1 adds three inert regression assertions: atomic replacement after path loading, repeated same-session validation and missing-resource refusal. They run with the existing eight setup-policy assertions during SourceOnly and full validation. These 11 assertions and PowerShell parsing/diff checks pass in the permitted agent PowerShell host. A Windows PowerShell 5.1 test attempt was blocked by AuthorizationManager in the AppContainer; no execution-policy override was used and 5.1 desktop execution remains to be verified by the owner. The agent's signature trust view also remains distinct from the approved desktop user's certificate view.

The corrected reader independently matches the actual version-14 embedded payload to its unchanged receipt. This is payload-read evidence, not a current-source freshness pass: packaging source has now changed and the installer must be rebuilt normally before full Test-EasySetup succeeds. The agent did not publish an unsigned replacement, alter certificate trust or execute installation. Rebuild with the existing approved signer in normal desktop PowerShell, then run Test-EasySetup.ps1. The new reader supports repeated validation in that same session without retaining stale path-loaded assembly content.
