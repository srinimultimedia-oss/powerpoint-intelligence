# AI data flow and context boundary

## Verified data flow

User-entered endpoint/deployment/mode → ServiceConnectionSettings → strict versioned XML → per-user local file. Load flows in reverse through byte/character bounds, no-DTD parsing, namespace/version/field validation. `apps/powerpoint-vsto/ServiceConnections.cs:58` `apps/powerpoint-vsto/ServiceConnections.cs:75`.

The stored profile path is `%LOCALAPPDATA%/PowerPointIntelligence/service-connections.v1.xml`; it is deliberately a compatibility identity, despite the NexWave display brand. No presentation, worksheet values, prompt text, token, credential or response field is in this schema. `apps/powerpoint-vsto/ServiceConnections.cs:35`.

## Not present

No presentation-context extraction for AI, upload, embedding store, vector search, model response persistence, chat history, retention scheduler, token budget or usage telemetry is implemented. PIIDX extracts local presentation metadata for a local library index, not an AI retrieval service. Do not join these independent systems into an invented RAG pipeline.

## Privacy limits

The profile is plaintext nonsecret metadata, not encrypted secret storage. Endpoint path components and deployment identifiers can still reveal organizational information. The UI warns that only approved nonsecret identifiers belong here. Future data minimization, retention, authorization and Azure processing policies are blocked by Q-005; see [AI security](ai-security.md).
