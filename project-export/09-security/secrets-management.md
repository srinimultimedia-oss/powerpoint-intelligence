# Secrets management and sensitive data

## Implemented boundary

Service settings hold endpoints/deployment/mode only, reject URI userinfo/query/fragment and unknown XML fields; tests reject ApiKey fields. No token/password value exists in the format. `apps/powerpoint-vsto/ServiceConnections.cs:45` `tests/s0/ServiceConnectionTests.cs:29`.

Signing private keys remain in Windows certificate storage. Builder selects an approved thumbprint with private key, code-signing EKU and trust chain; payload comprises DLL/config/manifest/VSTO/templates rather than an exported PFX. `scripts/Build-Integrity.ps1:57` `easySetup/Build-EasySetup.ps1:41`.

**NOT IMPLEMENTED:** Azure Key Vault, Windows Credential Manager connector, managed identity, runtime credential acquisition, encrypted application secret persistence or rotation. RuntimeCredential is only an enum.

## Sensitive local data

Presentation associations can retain source paths/worksheet/range/header signatures. PIIDX stores local presentation text/object metadata. These are user data despite no network upload. GetBuildInfo exposes assembly location; installation/repair evidence can retain local paths. Exported documentation copies no user documents or private credentials.

PowerPointHost diagnostics use fixed operation/error/HRESULT rather than workbook values. Scanner collector exports matching diagnostic records but explicitly makes no protection/clearance verdict. `apps/powerpoint-vsto/PowerPointHost.cs:54` `scripts/Collect-SceneThreatEvidence.ps1:60`.

## Recommendations

Use organization-approved publisher/renewal policy; never bundle keys or place secrets in endpoint paths. Define workstation catalog/profile ACL, retention/backup and future AI upload classification before production deployment. Q-005/Q-007 remain owner/environment decisions. Metadata is plaintext; absence of credential fields is not encrypted storage.
