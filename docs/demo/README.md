# Demonstration build

Owner requested broad local demonstration functionality before individual runtime
testing on 22 October 2026. External services are explicitly not configured.
This is an independent implementation, not a clone or a compatibility claim.

## Install and present

1. Save your work and close PowerPoint normally.
2. Run `scripts/Install-Local.ps1` to build/sign/install the changed add-in.
   Do not use `-SkipBuild` for source updates.
3. Open a disposable presentation with a blank slide.
4. Use **Intelligence → Create Demo Slides** to generate native samples in a NEW
   presentation. This command creates fifteen slides and can take time while
   embedded workbooks are initialized. Save the new file yourself.
5. Use **Elements** for chart/diagram insertion and **Explore Demo Tools** for
   the searchable tool catalog. Select slide objects before arrangement actions.
6. Check **Show full requested inventory** to inspect each imported feature's
   actual demonstration scope and status. The inventory is not an acceptance pass.

## Local behavior and limitations

- Native charts: columns/bars (clustered, stacked, 100%), line, row-oriented line
  comparison, stacked area, pie, doughnut, scatter, bubble and column/line combo.
  Sample data is clearly identified and editable in embedded Excel workbooks.
- Mekko, waterfall, Gantt, process flow and indicator commands create editable
  sample shapes. They are static diagrams, not full data-driven chart engines.
  No automatic label collision algorithm or Gantt scheduling engine is present.
- Shape actions include alignment, matching size, distribution, swap, scaling,
  grid arrangement, rotate/flip, duplicate, group, ungroup and z-order.
- Text tools include selected-object paste, font/size, decimal separators,
  proofing language and Unicode symbols. Text replacement can reset run formatting.
- Tables are native objects; the simple grid editor supports up to 400 cells.
- Agenda, image, template and slide import/export commands use local files.
  Library browsing is a file picker, not an enterprise content management system.
  Selected-slide export uses native clipboard transfer and does not certify
  managed chart XML preservation. No email is sent.
- Local XLSX linking reads the first worksheet used range, capped at 50x15 cells,
  numeric data plus text headers/categories. External-link updates are disabled
  while opening the read-only source. Store paths in slide CustomerData and use
  explicit manual refresh. Moved/missing files are not silently rebound. Original
  file paths remain in the presentation. This is a demonstration convention, not
  production approval of Q-010 semantics. Imported workbook formulas become values.
- Native chart options include title, value-axis title, labels, legend, data table
  and style. Static reference/difference annotations are not chart-axis linked.
- CAGR accepts entered values; workbook rounding preserves formulas and has a
  1,000-cell limit. Explicit comment removal covers legacy comments only and
  requires confirmation. It is not complete presentation sanitization.
- `sample-chart.s0.json` is the S0 model, NOT the S1 canonical v2 schema. JSON
  import validates that model and inserts a new managed clustered-column chart.
  Export works only on managed S0 charts. General native demo charts/diagrams do
  not acquire S0 canonical metadata or data-sync semantics.
- `sample-style.json` uses integer Office RGB/OLE colors and a font/size; it applies
  to the selection and is not a PowerPoint theme or proprietary style format.
- No proprietary Mekko Graphics file converter is implemented. Native slide import
  can copy existing editable objects but does not translate proprietary metadata.
- AI, Tableau, MCP, online research, chat history, organizational identity and
  enterprise deployment services are not configured and return no fake results.

## Experimental local automation API

After the add-in is loaded, its `COMAddIns["PowerPointIntelligence"].Object`
exposes `IDemoAutomation` (exact registered ProgID must be verified on the desktop):

| Method | Input | Output / behavior |
|---|---|---|
| GetVersion | none | Demo API/model version string |
| InsertChartJson | bounded S0 ChartDocument JSON | validates, inserts a new native chart, returns new GUID |

Calls must arrive on the Office UI STA; no server automation or network API is
provided. Unsupported model versions and invalid dimensions are rejected. This
is not the reference product's API, and no API compatibility is claimed.

## Validation status

Compilation and structural checks are the initial build checks. Individual demo
commands, UI layout, input failure recovery, keyboard navigation, Office build
matrix and save/reopen of the new tools are NOT runtime-certified. Existing S0
intermittent disconnection remains unresolved. The earlier two successful S0
adapter runs do not certify this substantially expanded build. Deferred tests
must be completed before production use or a claim of feature completeness.

The original inventory has 103 imported navigation-level rows; these are not 103
fully implemented or tested features. `demo-feature-inventory.json` records each
row's limited demo scope, missing behavior and configuration dependencies.
