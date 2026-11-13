# Azure AI integration

**VERIFIED (source):** The running add-in is local VSTO/COM software. Azure connectivity, AI inference and enterprise identity enforcement are not implemented. Service profiles are preparation metadata only. Runtime UI tests were not executed for this documentation export.

## Profile contract

`apps/powerpoint-vsto/ServiceConnections.cs:12` and `apps/powerpoint-vsto/ServiceConnections.cs:45` define the real configuration interface.

| Field | Constraint |
|---|---|
| Service | Exactly one each of AzureOpenAI, OnPremDataService, AssetProvider in version 1 |
| Endpoint | Empty or absolute HTTPS URI; maximum 2048 characters; no whitespace/control, userinfo, query or fragment |
| Deployment | Maximum 128 characters, trimmed, no controls; nonempty only allowed for AzureOpenAI |
| Authentication | Unselected, MicrosoftEntraId or RuntimeCredential; WindowsIntegrated additionally allowed only for OnPremDataService |
| Readiness label | Azure requires endpoint, deployment and selected mode; other services require endpoint/mode |

Selecting MicrosoftEntraId does **not** acquire a token; RuntimeCredential does **not** store a credential; WindowsIntegrated does **not** negotiate Windows authentication. The code has no Azure endpoint call, Azure API version, model-name resolution, client retry, token accounting, content filtering or billing implementation.

## Decisions and prerequisites

Q-003, Q-004 and Q-005 keep Azure region/deployment/egress, identity provider and data policy unresolved. AD-017 authorizes live service scope with separate configuration but does not close those decisions. `docs/DECISIONS.md:8` `docs/DECISIONS.md:8`.

A future connector must introduce a real authenticated provider interface, reviewed request schema, cancellation/timeouts, bounded responses, least-privilege identity and explicit document approval. These are recommendations, not implemented controls. No endpoint-specific OpenAPI description is generated because no actual Azure request contract exists here.

`tests/s0/ServiceConnectionTests.cs` verifies metadata/codec behavior with example.invalid addresses and asserts no network/UI call. It is not a live Azure integration test. `tests/s0/ServiceConnectionTests.cs:35`.
