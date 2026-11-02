# Typed native series and workbook read-back — 29 October 2026

Owner native evidence: artifacts/demo-features-b248b9960b9c4fceaa1b6806ba987c24. Context properties 23, font 8, productivity 13, text elements 14, process flow 18, indicators 16, connectors 20, chapter agenda 30 and workflow 34 native assertions pass. Mapped Excel stops on a non-zero-based SAFEARRAY cast. The demo chart report still reports zero points, followed by an RPC failure; only two catalog commands ran and 88 were blocked. Broad catalog acceptance is not complete.

## Changes

The remaining mapped-Excel test used a dynamic COM property result as an Array. The exception identifies an attempted conversion of System.Object[*] to System.Object[]. Native chart operations and assertions now call PowerPointHost.NativeSeriesValues: obtain an explicitly typed PowerPoint.Series, read Values into object, then normalize through the existing SeriesVector routine. No raw SAFEARRAY is cast to object[]. Series creation and the remaining chart-control Item paths also use explicit PowerPoint.Series declarations. The S0 capture path already used this typed pattern.

LinkedChartData.Write now verifies the actual embedded worksheet Value2 matrix immediately after assignment, before calculation/source binding. Dimensions, headers/categories, numeric types, signed values, zero and missing values must agree with the requested table. Native numeric cells cannot pass as text. Failures report row/column/type only and never cell contents. Four pure tests verify one-based read-back and refusal of wrong dimensions, value mismatch and numeric text.

This isolates the handoff: if worksheet read-back passes but chart points remain absent, the failure remains at the native chart binding/cache stage; it is not silently replaced with literal sample data or accepted on timeout. No series-formatting rebuild, fallback model or longer chart timeout was introduced.

Marker: PPI-DEMO-RIBBON-20261029-v17-piidx7-typed-series. Logs: artifacts/native-series-20261029. Native array normalization, linked workbook edits, chart readback and full catalog execution still require desktop validation with this new payload. Earlier owner desktop passes remain attributed to their original payload.
Final non-interactive regression: 23/23 PASS, including AnyCPU/x86/x64 production compilation. Evidence: artifacts/validation-20261029-222030-2b19bf2d/results.json; source snapshot B1FF642A266BA0A845AC63B544E9AE04358E1941475BC68271427556C9B24836. Interactive Office was not run in this agent session.
