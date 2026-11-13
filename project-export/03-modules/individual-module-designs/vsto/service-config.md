# Nonsecret external-service configuration UI

**VERIFIED — source implementation. Current runtime acceptance is not established by this design review.**

## Purpose and implementation status

Persist endpoint/deployment/authentication intent for future services without claiming working connectivity.

Implemented local source is distinguished from supported subset/absent external behavior below. See current test evidence in the export testing chapter; do not treat historical logs or method existence as a current runtime pass.

## Source locations and internal components

- `apps/powerpoint-vsto/ServiceConnections.cs:1` — ConnectionService, ConnectionAuthentication, ServiceConnectionProfile, ServiceConnectionSettings, ServiceConnections.

## Inputs, outputs, data structures and dependencies

Inputs: nonsecret endpoint/deployment/auth choice. Outputs: version1 local XML/status. Dependencies: PathPolicy.ConfigFile, System.Xml and WinForms; no HTTP client, authentication SDK or credential store.

## Internal execution and external integration

Exactly three profiles are supported: AzureOpenAI, OnPremDataService and AssetProvider. Validator requires HTTPS URI without credentials/query/fragment, bounded deployment only for Azure and allowed auth enums. Form reports Not configured or Configured, not connectivity-tested. XML Load returns defaults if absent; Save writes bounded UTF8 temp file and atomically replaces existing config.

## Configuration, validation and failure behavior

Endpoint<=2048,deployment<=128,XML<=16384 chars/file65536bytes; user LocalAppData/PowerPointIntelligence/service-connections.v1.xml. DTD/resolver disabled; malformed load warns replacement-on-save and Cancel preserves prior data.

## Security considerations and cleanup

Local native calls run in the current Office user process and require supported selection/ownership; a mutex/selection key is not authentication or RBAC. File/metadata content must remain bounded and error text must not expose data or credentials. Explicit owned-resource cleanup and error/rollback paths are part of the cited implementation, not a guarantee that a failed COM server can always be restored.

## Test references

Lexical module-class matches in test source (not instrumentation coverage; no test is claimed executed here):

- `tests/s0/ServiceConnectionTests.cs:19`.

## Known limitations and technical debt

Live Azure/data/assets,token acquisition,Windows auth negotiation,prompt orchestration,retries/usage/billing and secret storage are NOT IMPLEMENTED here. An auth enum is not authentication.

## Extension points

Add authenticated typed connectors and reviewed credential/privacy controls; never equate saved profile completeness with tested service availability.

## Classes, interfaces and method navigation

Method names below are source declarations; private/internal members are not public API promises. The source hash/review manifest gives exact reviewed bytes.

| Declaration | Kind | Source |
|---|---|---|
| `ConnectionService` | enum | `apps/powerpoint-vsto/ServiceConnections.cs:12` |
| `ConnectionAuthentication` | enum | `apps/powerpoint-vsto/ServiceConnections.cs:13` |
| `ServiceConnectionProfile` | class | `apps/powerpoint-vsto/ServiceConnections.cs:14` |
| `Copy` | method | `apps/powerpoint-vsto/ServiceConnections.cs:20` |
| `ServiceConnectionSettings` | class | `apps/powerpoint-vsto/ServiceConnections.cs:22` |
| `Empty` | method | `apps/powerpoint-vsto/ServiceConnections.cs:26` |
| `ServiceConnections` | class | `apps/powerpoint-vsto/ServiceConnections.cs:31` |
| `AllowedAuthentication` | method | `apps/powerpoint-vsto/ServiceConnections.cs:36` |
| `Validate` | method | `apps/powerpoint-vsto/ServiceConnections.cs:38` |
| `ValidateProfile` | method | `apps/powerpoint-vsto/ServiceConnections.cs:45` |
| `Status` | method | `apps/powerpoint-vsto/ServiceConnections.cs:53` |
| `Encode` | method | `apps/powerpoint-vsto/ServiceConnections.cs:58` |
| `Attributes` | method | `apps/powerpoint-vsto/ServiceConnections.cs:67` |
| `Elements` | method | `apps/powerpoint-vsto/ServiceConnections.cs:71` |
| `Decode` | method | `apps/powerpoint-vsto/ServiceConnections.cs:75` |
| `Load` | method | `apps/powerpoint-vsto/ServiceConnections.cs:89` |
| `Save` | method | `apps/powerpoint-vsto/ServiceConnections.cs:93` |
| `ShowConfiguration` | method | `apps/powerpoint-vsto/ServiceConnections.cs:98` |
| `Label` | method | `apps/powerpoint-vsto/ServiceConnections.cs:119` |
