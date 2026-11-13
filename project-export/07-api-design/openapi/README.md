# OpenAPI applicability

**NOT IMPLEMENTED:** No actual HTTP API contract can be established. The gateway directory contains only `services/ai-gateway/README.md:1` and `services/ai-gateway/action-plan.schema.json:1`. ServiceConnections stores endpoint preparation metadata without a client. There is no established route, HTTP method, auth protocol, status-code mapping or request/response exchange to serialize as OpenAPI.

A speculative OpenAPI document would misrepresent planned architecture as implemented. The real COM, managed, native and file contracts are documented in the parent API documents. Create an OpenAPI document only after the gateway implementation and contract tests exist.
