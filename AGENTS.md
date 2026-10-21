# Codex instructions

## Mission
Implement the independent PowerPoint add-in specified in `docs/requirements/architecture-functional-design.docx`. Work incrementally, test-first, and never claim functionality not verified.

## Fixed decisions
Windows desktop first, on-prem hosting, Azure OpenAI provider, native editable PowerPoint elements, detailed engineering deliverables.

## Non-negotiable engineering rules
1. **Read `docs/DECISIONS.md` before making architecture changes.** For BLOCKED decisions, create a documented spike and request clarification; do not silently choose.
2. The VSTO in-process add-in targets **.NET Framework 4.8**, not .NET 8. The on-prem API services may use current supported .NET.
3. Do not automate Office on an unattended server. Server-side PPTX generation uses Open XML or an appropriate document generation engine.
4. Never reverse engineer or copy think-cell binaries, proprietary assets, or protected implementation details.
5. Use typed DTOs, schema validation, and explicit versioned persistence. AI output must not directly call COM or mutate documents.
6. Never place Azure keys, client secrets, or workbook contents in logs or git.
7. Each change needs tests, reproduction steps, and explicit limitations. Prefer small PR-sized tasks.
8. Use native editable objects where feasible; record exceptions and demonstrate PowerPoint save/reopen.

## Initial execution order
- S0-A: Windows environment and Office/VSTO compatibility report.
- S0-B: VSTO ribbon and task pane bootstrap.
- S0-C: Insert native clustered-column chart or editable shapes via COM.
- S0-D: Persist canonical chart metadata and restore after reopen.
- S0-E: Run round-trip, duplicate-slide, and error recovery tests.
- S1: Add data model, command bus, chart rendering, and tests.

## Definition of done
No acceptance test marked passed without execution evidence (machine/Office version, steps, result, screenshot or test log). Update requirements traceability for each feature.
