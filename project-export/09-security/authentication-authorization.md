# Authentication and authorization

## Actual identities

The add-in executes under the desktop Windows user and Office process. VSTO loading uses publisher trust. COMAddIns.Object exposes local COM automation, not bearer-token HTTP calls. No application login/user/role table, Entra token acquisition, Windows Negotiate handshake or identity-provider SDK exists in authored runtime/service sources.

ServiceConnectionProfile.Authentication selects future intent; it never authenticates a session. Status expressly distinguishes metadata readiness from connectivity testing. `apps/powerpoint-vsto/ServiceConnections.cs:13` `apps/powerpoint-vsto/ServiceConnections.cs:56`.

## Guards that are not RBAC

Office compatibility enables commands based on Office 16 API/process architecture/minimum build. COM mutations require creating UI STA. Contextual edits verify selection snapshots before applying a queued change. OfficeMutationLease coordinates writers across tests/add-in timers; named mutex ownership is not a permission grant. Filesystem access follows Windows ACL/host policy; the add-in does not elevate itself.

Evidence: `apps/powerpoint-vsto/OfficeCompatibility.cs:61` `apps/powerpoint-vsto/DemoApi.cs:35` `apps/powerpoint-vsto/PowerPointHost.cs:14`.

## Planned boundary

Gateway README proposes RBAC before authorized context retrieval. Q-004 leaves identity provider, SSO protocol, groups and service accounts unresolved. This is **PLANNED / NOT IMPLEMENTED**, not an existing enterprise identity design. `services/ai-gateway/README.md:5` `docs/DECISIONS.md:8`.

Future implementation should validate issuer/audience/scopes, map permissions to specific actions/documents, reject unapproved targets/revisions and audit denials. These are recommendations; no actual role names or token grants can be inferred.
