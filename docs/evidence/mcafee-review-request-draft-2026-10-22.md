# Draft McAfee detection review request — not submitted

Please review a suspected false-positive detection affecting a locally compiled .NET Framework 4.8 x64 scene-contract test runner. We are requesting a verdict and explanation; we are not asserting that the detection is incorrect.

Product observed: McAfee 16.3.1.108 on Windows 11 x64. Affected filename: SceneContractTests.exe, under a PowerPoint add-in repository's artifacts/scene-contracts-x64 directory and immutable validation snapshots.

Confirmed matching records:

| Time UTC on 22 October 2026 | SHA256 | Detection | Action |
|---|---|---|---|
| 12:03:34 | EA5C34B16761FD30AC9ADF0F8CB4DE362B1AA81A93DD9D170D06EA0A087FA7A3 | ti!EA5C34B16761 | infection quarantined |
| 12:04:27 | 72BE5A0B44966754BC2D7911861419415A40489490314A420B265F3800AE7F1E | ti!72BE5A0B4496 | infection quarantined |

Both records report sensor `section execute` and final detection source `rp-s`. What evidence or rule triggered these verdicts, and are these builds classified correctly?

Build provenance: Visual Studio Community Roslyn csc.exe with a valid Microsoft Corporation Authenticode signature; .NET Framework 4.8 reference assemblies; Newtonsoft.Json 13.0.3 dependency with locally validated Authenticode signature; local development test EXE/DLL unsigned. Compile target /platform:x64, ordinary managed executable, no packing or obfuscation. Source review found fixture reads, whitelisted DTO validation, SHA256 scene hashing, and JSON round-trip writes; no process launch, network retrieval, registry modification, dynamic assembly loading or self-deletion in the runner/contract sources. These observations do not independently prove a malware-free artifact.

The user paused protection for an initial retry, then re-enabled it. The same subsequently built hash E02CDBF2FEFB13AD09CB45D78750391FDC49A38AE46541A29D94A4E6BD96B17E then passed while protection was active. A fresh protected full-suite build, hash F141AB9796AD09C6E0B63CD3B1DE321C2394BB4C7C222F954DFA076FCAB47990, also passed. No additional matching quarantine was recorded by the final log read. This current behavior does not establish the cause or resolution of earlier verdicts.

Please advise whether the earlier builds require sample review and whether a detection update is needed. No executable, source file, customer presentation, account identifier, private path or full telemetry log has been uploaded with this draft. Artifact transfer requires explicit owner approval. The quarantined builds have not been exported or restored to obtain samples.

Official review procedure: https://www.mcafee.com/support/s/article/000001921
