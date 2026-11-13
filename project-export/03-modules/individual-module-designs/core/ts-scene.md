# TypeScript scene construction and v2 file adapter

<!-- generated-toc -->
## Contents

- [1. Purpose](#1-purpose)
- [2. Implementation status](#2-implementation-status)
- [3. Source locations](#3-source-locations)
- [4. Internal components](#4-internal-components)
- [5. Classes and interfaces](#5-classes-and-interfaces)
- [6. Important methods and functions](#6-important-methods-and-functions)
- [7. Inputs and outputs](#7-inputs-and-outputs)
- [8. Dependencies](#8-dependencies)
- [9. Workflow](#9-workflow)
- [10. Integration points](#10-integration-points)
- [11. Exceptions and recovery](#11-exceptions-and-recovery)
- [12. Configuration dependencies](#12-configuration-dependencies)
- [13. Security considerations](#13-security-considerations)
- [14. Unit and integration test definitions](#14-unit-and-integration-test-definitions)
- [15. Known limitations](#15-known-limitations)
- [16. Technical debt](#16-technical-debt)
- [17. Extension points](#17-extension-points)
<!-- end-generated-toc -->

**VERIFIED source review.** This design describes inspected code, not current runtime/Office execution.

## 1. Purpose

Constructs whitelisted sealed scenes from normalized data; local CLI adapts canonical v2 JSON.

## 2. Implementation status

Offline Node tooling, not REST/daemon or automatic in-process JavaScript execution by VSTO.

## 3. Source locations

- `packages/chart-engine/src/scene-builder.ts:1`
- `packages/chart-engine/src/scene-contract.ts:1`
- `packages/chart-engine/src/strict-scene-json.ts:1`
- `packages/chart-engine/tools/render-chart-v2.mjs:1`
- `packages/chart-domain/src/validate.mjs:1`
- `packages/chart-domain/package.json:1`

## 4. Internal components

Pure calculations, WebCrypto SHA-256, AJV2020 and Node fs only in schema-loader/CLI boundary. Validate request/palette cardinality; map bars/markers/paths/areas/sectors; snapshot before await; hash and verify strict JSON. V2 adapter fixes vertical clustered gap 0.2, normalizes GUID case, preserves revision and colors and emits AXES_NOT_RENDERED plus optional LEGEND_NOT_RENDERED.

## 5. Classes and interfaces

- `packages/chart-engine/src/scene-builder.ts:6` — `interface BaseRequest`
- `packages/chart-engine/src/scene-builder.ts:7` — `export type SceneRequest = BaseRequest & (`
- `packages/chart-engine/src/scene-contract.ts:6` — `export interface Point`
- `packages/chart-engine/src/scene-contract.ts:7` — `export type Geometry =`
- `packages/chart-engine/src/scene-contract.ts:12` — `export interface Paint`
- `packages/chart-engine/src/scene-contract.ts:13` — `export interface SceneNode`
- `packages/chart-engine/src/scene-contract.ts:14` — `export interface SceneDiagnostic`
- `packages/chart-engine/src/scene-contract.ts:15` — `export interface ScenePayload`
- `packages/chart-engine/src/scene-contract.ts:19` — `export interface RenderResult extends ScenePayload`

## 6. Important methods and functions

Declaration locator below is lexical and may include overload/call-like lines; semantic workflow is described separately.

- `packages/chart-engine/src/scene-builder.ts:13` — `export async function buildScene(request: SceneRequest): Promise<RenderResult>`
- `packages/chart-engine/src/scene-contract.ts:43` — `export function validateScene(value: unknown, sealed?: true): asserts value is RenderResult;`
- `packages/chart-engine/src/scene-contract.ts:44` — `export function validateScene(value: unknown, sealed: false): asserts value is ScenePayload;`
- `packages/chart-engine/src/scene-contract.ts:45` — `export function validateScene(value: unknown, sealed = true): asserts value is ScenePayload`
- `packages/chart-engine/src/scene-contract.ts:92` — `export function sceneHashBytes(payload: ScenePayload): Uint8Array`
- `packages/chart-engine/src/scene-contract.ts:130` — `export async function sealScene(payload: ScenePayload): Promise<RenderResult>`
- `packages/chart-engine/src/scene-contract.ts:136` — `export async function verifyScene(value: unknown): Promise<RenderResult>`
- `packages/chart-engine/src/scene-contract.ts:143` — `export async function decodeScene(json: string): Promise<RenderResult>`
- `packages/chart-engine/src/scene-contract.ts:144` — `export async function encodeScene(scene: RenderResult): Promise<string>`
- `packages/chart-engine/src/strict-scene-json.ts:5` — `export function parseSceneJson(json: string): unknown`
- `packages/chart-engine/tools/render-chart-v2.mjs:9` — `export async function renderChartV2(json)`
- `packages/chart-domain/src/validate.mjs:7` — `export function validateChart(model)`

## 7. Inputs and outputs

Closed SceneRequest -> Promise RenderResult; v2 JSON -> clustered scene; CLI local file input/output.

## 8. Dependencies

Pure calculations, WebCrypto SHA-256, AJV2020 and Node fs only in schema-loader/CLI boundary.

## 9. Workflow

Validate request/palette cardinality; map bars/markers/paths/areas/sectors; snapshot before await; hash and verify strict JSON. V2 adapter fixes vertical clustered gap 0.2, normalizes GUID case, preserves revision and colors and emits AXES_NOT_RENDERED plus optional LEGEND_NOT_RENDERED.

## 10. Integration points

Closed SceneRequest -> Promise RenderResult; v2 JSON -> clustered scene; CLI local file input/output. Offline Node tooling, not REST/daemon or automatic in-process JavaScript execution by VSTO.

## 11. Exceptions and recovery

CalculationError scene.*; CLI sanitized catch sets process exit code 1; no automatic retry.

## 12. Configuration dependencies

Scene v1 / pt / 8 MiB, canonical v2 1 MiB; palette count series or pie categories; explicit orientation, gap policy and angle.

## 13. Security considerations

Closed properties, valid Unicode, null-prototype parser maps, no executable nodes. Hash integrity does not authorize COM.

## 14. Unit and integration test definitions

- `packages/chart-engine/test/scene-contract.test.mjs:1`
- `packages/chart-engine/test/canonical-scene.test.mjs:1`
- `packages/chart-engine/test/scene-schema.test.mjs:1`
- `tests/s3/validate-roundtrip.mjs:1`

Inspected definitions only. See the parent current test report for execution; historical passes do not establish current acceptance.

## 15. Known limitations

V2 projection has no axes/legend and only one canonical family; v3 is not CLI input; broader scene DTO != native capability.

## 16. Technical debt

C#/TS canonical hash and contract changes require cross-language governance.

## 17. Extension points

Extend request/schema/codec fixtures together, retain separate execution authorization boundary.
