# OPS-01 — Service profile configuration and AI scaffold

**Evidence category:** source verified; execution status is stated separately. Baseline includes the actual inspected checkout, not an assumption that every working byte equals HEAD.

## Design

| Dimension | Implementation-grounded description |
|---|---|
| 1. Purpose | Prepare nonsecret local connection metadata without pretending to have a live provider. |
| 2. Implementation status | Profile validation/codec/file/UI implemented; AI gateway/identity/inference NOT IMPLEMENTED. Gateway schema is PLANNED. |
| 3. Internal components/classes/interfaces | ConnectionService, ConnectionAuthentication, ServiceConnectionProfile, ServiceConnectionSettings and ServiceConnections; AIActionPlan schema is not a runnable class. |
| 4. Important functions/methods | Empty, AllowedAuthentication, Validate/ValidateProfile, Status, Encode/Decode, Load/Save and ShowConfiguration. |
| 5. Inputs and outputs | Three bounded service profiles, version1 strict XML; outputs readiness label, local XML file or validation error. No token or AI response. |
| 6. Dependencies | WinForms, XmlReader/XmlDocument, LINQ, UTF8Encoding and Library.Core PathPolicy. ThisAddIn connections workflow invokes settings. |
| 7. Internal workflow | Load existing/default settings; show warnings if unreadable; edit endpoint/mode/deployment; live syntax status; Cancel preserves old file; Save validates, writes sibling temporary XML and atomically moves/replaces target. |
| 8. External integrations | Only local filesystem and user dialog; no HTTP client/authentication/credential store. |
| 9. Exceptions and error handling | InvalidDataException/XmlException/IO/Unauthorized/decoder errors produce safe defaults with warning or failed save. A malformed prior file is replaced only on explicit valid Save. |
| 10. Configuration | LocalAppData/PowerPointIntelligence/service-connections.v1.xml; version1, max16384 XML characters/max65536 saved bytes; exactly3 profiles. |
| 11. Security | HTTPS URI syntax, no credentials/query/fragment/whitespace, known enums/fields, no DTD; schema has no secret values. Any HTTPS host is not a provider allowlist. |
| 12. Tests | ServiceConnectionTests covers empty status, constraints, malicious/nested/unknown XML, byte bounds, atomic replace and invalid-save preservation. No network/UI operation asserted. |
| 13. Known limitations | Selecting Entra/WindowsIntegrated/RuntimeCredential is intent only; status means not connectivity-tested. No prompts/history/tool calls/retries/usage metrics exist. |
| 14. Technical debt | Gateway generic arguments and approval flags need typed semantics; identity/egress/data policy decisions remain unresolved. |
| 15. Extension points | Add an explicit authenticated provider abstraction with cancellation, bounded typed request/response, target/revision approval and audit before connecting settings to live calls. |

## 16. Complete owned source map

| Source | Analysis | Responsibility/declarations |
|---|---|---|
| `apps/powerpoint-vsto/ServiceConnections.cs:1` | Complete text read; no runtime execution | Declarations: Copy, Empty, AllowedAuthentication, Validate, ValidateProfile |
| `services/ai-gateway/README.md:1` | Complete text read; no runtime execution | Script/project entry; its actual flow and responsibilities are documented above. |
| `services/ai-gateway/action-plan.schema.json:1` | Complete text read; no runtime execution | Script/project entry; its actual flow and responsibilities are documented above. |
| `tests/s0/ServiceConnectionTests.cs:1` | Complete text read; no runtime execution | Declarations: Check, Reject, Valid, Run |

## 17. Handover and maintenance checks

Review the exact source/hash/line counts in `09-security/reviewed-files.csv` and this directory’s `source-module-mapping.csv`. Do not substitute historical PASS logs for current execution. Preserve unrelated module behavior and tests before modifying a boundary; new methods must have typed inputs, explicit error semantics and a corresponding negative test. See the dedicated AI/API/security/deployment documents for this module’s external contract.
