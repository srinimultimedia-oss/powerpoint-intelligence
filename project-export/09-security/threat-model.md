# Threat model

Scope: local VSTO runtime, document/data/index files, native Office/Windows calls, development scripts and signed deployment. Assets: native document integrity/editability, local source data/catalog text and build/installation provenance. Threat actors include malformed input authors, corrupt/stale metadata, local file replacement, compromised dependencies and future overprivileged connectors. A fully privileged workstation administrator is outside the application's own boundary.

| Scenario | Existing control | Residual / verification |
|---|---|---|
| Traversal/ADS escape from generated path | PathPolicy syntax/Within, owned-output ancestor checks, Artifact-Inputs | User-selected inputs intentionally have broader authority; junction tests limited by host |
| XML external entity/resource exhaustion | Explicit bounded no-DTD/resolver-null readers in codecs and Safe-Xml | Install-Local/Test-DemoPackage implicit casts need consistent review |
| Ambiguous/oversize JSON | Strict bounded grammar/depth32/duplicate-key/finite-number and semantic codecs | S0 COM DataContract JSON differs from strict v2/scene path |
| Corrupt copied native metadata | Version/identity/namespace checks, explicit mapping and safe preservation on failure | Native family/copy/reopen acceptance requires Office execution |
| Cached executable plus forged receipt | Independently reproduce current source; lock verified bytes/files | Source/compiler/build user trusted; receipt alone insufficient |
| Compromised development dependency | JS lockfile/lifecycle scripts disabled; scene-specific publisher verification | Test-S1's initial NuGet download lacks pinned checksum before use |
| Malicious native source template | Bounded ZIP/XML extraction and intended read-only source access | Hidden/read-only Office source is not a sandbox |
| Tampered install package | Authenticode, bundle/file hashes, extraction bounds/SafePath | Target estate publisher trust remains separate |
| Future prompt injection/exfiltration | AI currently disconnected/no executor | Typed actions/auth/approval/minimization required before connection |

Evidence: `scripts/Artifact-Inputs.ps1:1` `scripts/Safe-Xml.ps1:1` `scripts/Test-S1.ps1:18` `src/Shared.Contracts/BoundedJson.cs:28` `apps/powerpoint-vsto/DemoApi.cs:40` `easySetup/Bootstrapper.cs:53`.

No exploitation/network probes were executed. Gap severity is engineering assessment, not a new scanner severity or exploit confirmation.
