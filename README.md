# PowerPoint Intelligence — Codex Starter

**Status:** S0 implementation under development; Windows VSTO runtime acceptance is unverified. See [S0 evidence](docs/evidence/S0-2026-10-21.md) and [validation runbook](docs/implementation/S0-windows-validation.md). Do not proceed to S1 until all S0 acceptance criteria pass in actual PowerPoint.

## Approved baseline
- Windows desktop PowerPoint first; C# VSTO/COM adapter (.NET Framework 4.8)
- On-premises .NET 8+ application services; Azure OpenAI through controlled gateway (not an on-prem model)
- Native editable PowerPoint shapes/charts as feasible, plus versioned canonical metadata
- Full functional design: `docs/requirements/architecture-functional-design.docx`
- AI must propose structured commands; deterministic calculation/validation executes changes

## Quick start for Codex
1. Read `AGENTS.md` and the Word requirements document.
2. Review `docs/DECISIONS.md`; do not invent unknown Office versions, tenant IDs, endpoint secrets or licensing requirements.
3. Start with `docs/implementation/S0-feasibility.md`. Build the Windows VSTO project in Visual Studio.
4. Implement the host interface and complete insert → save → reopen → inspect round-trip.
5. Run chart engine tests: `cd packages/chart-core && npm install && npm test` (Node.js 20+).
6. Do not start Azure calls until networking, endpoint, identity and data-handling decisions are approved.

## Repository map
- `apps/powerpoint-vsto/`: .NET Framework 4.8 VSTO project, ribbon/status pane, native chart adapter and bounded versioned XML persistence (full VSTO build/load unverified)
- `tests/s0/`, `scripts/Test-S0.ps1`: metadata tests and opt-in interactive Office integration harness
- `packages/chart-core/`: deterministic chart calculations with unit tests
- `services/ai-gateway/`: on-prem AI action contract and policy placeholders
- `docs/`: decisions, development milestones, architecture and source SRS
- `.github/workflows/`: portable chart-core CI

## First release gate
A user can insert a chart into PowerPoint, save PPTX, close/reopen it, and recover editable chart metadata with correct object associations. No claim of completion until tested on Windows PowerPoint.
