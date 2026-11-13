# Prompt and action orchestration

## Actual implementation

**NOT IMPLEMENTED:** There is no prompt store, system/user prompt builder, conversation orchestration, agent runtime, retrieval pipeline, function dispatcher from model output or approval service. The only AI artifact is `AIActionPlan` JSON Schema, plus local profile UI. `services/ai-gateway/action-plan.schema.json:1` `services/ai-gateway/README.md:5`.

## Planned action schema

The schema allows status `proposed`, `validated`, `approved`, `executed`, `failed`; required root fields are planId, status, actions and approvalRequired. Action types are CREATE_CHART, UPDATE_CHART, CONVERT_CHART, APPLY_LAYOUT, ADD_ANNOTATION, GENERATE_SLIDE and GENERATE_REPORT. Each action requires actionId/type/arguments and may carry targetId/expectedRevision. Unknown root/action properties are rejected, but arguments is an unrestricted object. `services/ai-gateway/action-plan.schema.json:5`.

**Important semantic gap:** the schema permits approvalRequired=false and does not enforce approval-state transitions, positive actionId length, nonempty actions or per-action argument constraints. No runtime schema validation of AI responses is wired to the add-in. The existing domain validators and native command safeguards cannot be described as an AI execution pipeline without that integration.

## Extension sequence (recommendation)

Before implementing execution, define typed arguments per allowed action; resolve target/revision within the active document; validate deterministically before mutation; require approved plans where policy demands; expose a preview; execute on the Office STA via bounded commands; record an audit event without document contents. Never accept arbitrary code, SQL or COM member names from generated text. These are future engineering requirements derived from the gateway README, not existing workflow states.
