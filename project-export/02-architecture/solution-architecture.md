# Solution and container architecture

The unit of desktop deployment is the signed VSTO manifest/payload, loaded by PowerPoint. C# source is compiled from explicit project items, including linked domain/rendering/library sources. There is no `.sln` in this baseline; the two `.csproj` files are the build entry points. A source folder's presence does not establish inclusion: V3 contracts are absent from both active C# project compile lists. `src/Chart.Domain/Chart.Domain.csproj:12`; `apps/powerpoint-vsto/PowerPointIntelligence.csproj:39`.

```mermaid
flowchart TB
  subgraph Desktop[Interactive Windows desktop]
    PPT[PowerPoint process] --> VSTO[VSTO lifecycle and ribbon]
    VSTO --> UI[WinForms task pane dialogs contextual controls]
    UI --> TOOLS[Command and editing modules]
    TOOLS --> HOST[PowerPoint COM adapter]
    TOOLS --> MODELS[Canonical codecs and validation]
    TOOLS --> LIB[Library catalog and native import]
    LIB --> SQL[Windows winsqlite3 SQLite]
    LIB --> CACHE[Local files and preview cache]
    HOST --> PPT
  end
  subgraph Dev[Build and validation process boundary]
    NODE[Node TypeScript calculations and scene contracts]
    PS[PowerShell build and test scripts]
    CSC[Visual Studio Roslyn and MSBuild]
    PS --> CSC
    PS --> NODE
  end
  NODE -. canonical JSON and scene JSON artifacts .-> MODELS
```

Node tooling validates deterministic calculations and emits a geometry-only scene when its CLI is invoked. The production VSTO adapter is not documented as launching Node on every chart command; C# supplies its own validated clustered projection. `packages/chart-engine/tools/render-chart-v2.mjs:9`; `src/Chart.Renderers/NativeRenderingCore.cs:18`.

State is split across PPTX custom XML/tags, linked local input files, local JSON/XML configuration, PIIDX SQLite and PNG previews, plus build/cache receipts. There is no deployed on-prem application database in the repository. Local PIIDX is implemented and is separate from the unresolved enterprise database choice. `src/Library.Core/PiidxSchema.cs:1`; `docs/DECISIONS.md:1`.
