# AI security

## Implemented controls

- Endpoint syntax requires HTTPS and rejects URI credentials, query and fragment; bounded text prevents unbounded profile input. `apps/powerpoint-vsto/ServiceConnections.cs:45`
- Strict XML codec rejects DTD/external resolution, unknown attributes/elements, nested endpoint data and unknown enum/version/namespace values. `apps/powerpoint-vsto/ServiceConnections.cs:75`
- No credential value fields, no auth attempt, no outbound client and no simulated successful AI response. `apps/powerpoint-vsto/ServiceConnections.cs:29`
- Local save validates before filesystem publication and atomically replaces a previous profile. `apps/powerpoint-vsto/ServiceConnections.cs:93`

## Missing controls because execution is absent

**NOT IMPLEMENTED:** token authentication, RBAC, prompt injection isolation, model tool authorization, PII reduction, output policy enforcement, egress allowlisting, content safety, usage rate limiting and AI audit trails. An HTTPS URI check alone does not establish a trusted Azure tenant/host or an egress allowlist. The allowed endpoint can be any syntactically valid HTTPS host.

The action schema is not an authorization boundary: arguments is generic and approvalRequired can be false. A later executor must validate capabilities/targets/revisions independently, not trust a model's status label. `services/ai-gateway/action-plan.schema.json:1`.

## Verification boundary

ServiceConnectionTests cover metadata boundaries and persistence with inert addresses, not network identity or tenant policy. No real Azure resource, credentials or customer content was accessed by this export. Open decision evidence: `docs/DECISIONS.md:8` `docs/DECISIONS.md:8`.
