# PowerPoint Intelligence — Codex Starter

**Status:** Project initialized as a source-code scaffold. **Not** deployed to GitHub, Codex, or PowerPoint. The VSTO add-in must be created/compiled/tested on a Windows workstation with supported Visual Studio, Office desktop, and .NET Framework tooling.

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
- `apps/powerpoint-vsto/`: Windows add-in contracts and implementation guide (not yet a compiled VSTO project)
- `packages/chart-core/`: deterministic chart calculations with unit tests
- `services/ai-gateway/`: on-prem AI action contract and policy placeholders
- `docs/`: decisions, development milestones, architecture and source SRS
- `.github/workflows/`: portable chart-core CI

## First release gate
A user can insert a chart into PowerPoint, save PPTX, close/reopen it, and recover editable chart metadata with correct object associations. No claim of completion until tested on Windows PowerPoint.
