# Universal connector geometry and endpoint bindings

**VERIFIED — source implementation. Current runtime acceptance is not established by this design review.**

## Purpose and implementation status

Create native straight/elbow/curve connectors with canonical endpoint/style state and explicit reflow.

Implemented local source is distinguished from supported subset/absent external behavior below. See current test evidence in the export testing chapter; do not treat historical logs or method existence as a current runtime pass.

## Source locations and internal components

- `apps/powerpoint-vsto/UniversalConnectors.cs:1` — ConnectorEndpoint, ConnectorSpec, ConnectorCodec, ConnectorBindingPlan, ConnectorGeometryPlan, UniversalConnectors, Saved, EndpointEditor.

## Inputs, outputs, data structures and dependencies

Inputs: endpoint shape/site/series/point/free coordinate, route/dash/arrows/color/weight. Outputs: native ConnectorFormat bindings and ConnectorCodec XML. Dependencies: PowerPoint chart point/shape geometry, pure binding/geometry plans, XmlPartUpdates.

## Internal execution and external integration

Create validates ConnectorSpec, resolves site/chart-point/free endpoints and slide bounds, inserts native connector, binds sites and attaches XML/tag. Live capture reconciles native connections with prior metadata and transforms native rotated/flipped geometry into slide coordinates. Reflow preflights all managed connectors, snapshots geometry/bindings/XML and applies/restores updates. Native rotation is retained during capture.

## Configuration, validation and failure behavior

Anchors center/top/right/bottom/left; site<=1024,series<=100,point<=10000; explicit allowed routes/styles and finite coordinates inside slide. Endpoint target identities must resolve uniquely.

## Security considerations and cleanup

Local native calls run in the current Office user process and require supported selection/ownership; a mutex/selection key is not authentication or RBAC. File/metadata content must remain bounded and error text must not expose data or credentials. Explicit owned-resource cleanup and error/rollback paths are part of the cited implementation, not a guarantee that a failed COM server can always be restored.

## Test references

Lexical module-class matches in test source (not instrumentation coverage; no test is claimed executed here):

- `tests/s0/ContextPropertiesOfficeTests.cs:72`.
- `tests/s0/UniversalConnectorOfficeTests.cs:31`.
- `tests/s0/UniversalConnectorTests.cs:23`.

## Known limitations and technical debt

Only owned connectors reflow. Unsupported geometry/missing targets/ambiguous metadata reject without guessing. Chart-point anchors depend on supported accessible chart geometry. Automatic obstacle avoidance and arbitrary routing constraints are not implemented. Compensating restoration can fail and requests Undo.

## Extension points

Extend pure geometry/binding plans before native adapters; test rotated/flipped attached endpoints and XML after save/reopen.

## Classes, interfaces and method navigation

Method names below are source declarations; private/internal members are not public API promises. The source hash/review manifest gives exact reviewed bytes.

| Declaration | Kind | Source |
|---|---|---|
| `ConnectorEndpoint` | class | `apps/powerpoint-vsto/UniversalConnectors.cs:14` |
| `Copy` | method | `apps/powerpoint-vsto/UniversalConnectors.cs:19` |
| `Validate` | method | `apps/powerpoint-vsto/UniversalConnectors.cs:20` |
| `ConnectorSpec` | class | `apps/powerpoint-vsto/UniversalConnectors.cs:22` |
| `Finite` | method | `apps/powerpoint-vsto/UniversalConnectors.cs:29` |
| `Validate` | method | `apps/powerpoint-vsto/UniversalConnectors.cs:30` |
| `Copy` | method | `apps/powerpoint-vsto/UniversalConnectors.cs:31` |
| `AnchorAt` | method | `apps/powerpoint-vsto/UniversalConnectors.cs:32` |
| `ConnectorCodec` | class | `apps/powerpoint-vsto/UniversalConnectors.cs:35` |
| `Encode` | method | `apps/powerpoint-vsto/UniversalConnectors.cs:39` |
| `Shape` | method | `apps/powerpoint-vsto/UniversalConnectors.cs:41` |
| `Decode` | method | `apps/powerpoint-vsto/UniversalConnectors.cs:42` |
| `ConnectorBindingPlan` | class | `apps/powerpoint-vsto/UniversalConnectors.cs:47` |
| `Reconcile` | method | `apps/powerpoint-vsto/UniversalConnectors.cs:49` |
| `ConnectorGeometryPlan` | class | `apps/powerpoint-vsto/UniversalConnectors.cs:60` |
| `Position` | method | `apps/powerpoint-vsto/UniversalConnectors.cs:62` |
| `UniversalConnectors` | class | `apps/powerpoint-vsto/UniversalConnectors.cs:70` |
| `Active` | method | `apps/powerpoint-vsto/UniversalConnectors.cs:72` |
| `IsManaged` | method | `apps/powerpoint-vsto/UniversalConnectors.cs:73` |
| `Read` | method | `apps/powerpoint-vsto/UniversalConnectors.cs:74` |
| `Find` | method | `apps/powerpoint-vsto/UniversalConnectors.cs:75` |
| `Locate` | method | `apps/powerpoint-vsto/UniversalConnectors.cs:76` |
| `Route` | method | `apps/powerpoint-vsto/UniversalConnectors.cs:81` |
| `Arrow` | method | `apps/powerpoint-vsto/UniversalConnectors.cs:82` |
| `Style` | method | `apps/powerpoint-vsto/UniversalConnectors.cs:83` |
| `Geometry` | method | `apps/powerpoint-vsto/UniversalConnectors.cs:84` |
| `Bounds` | method | `apps/powerpoint-vsto/UniversalConnectors.cs:89` |
| `Position` | method | `apps/powerpoint-vsto/UniversalConnectors.cs:90` |
| `Live` | method | `apps/powerpoint-vsto/UniversalConnectors.cs:95` |
| `Snapshot` | method | `apps/powerpoint-vsto/UniversalConnectors.cs:101` |
| `Part` | method | `apps/powerpoint-vsto/UniversalConnectors.cs:105` |
| `Saved` | class | `apps/powerpoint-vsto/UniversalConnectors.cs:106` |
| `Capture` | method | `apps/powerpoint-vsto/UniversalConnectors.cs:111` |
| `Restore` | method | `apps/powerpoint-vsto/UniversalConnectors.cs:112` |
| `Create` | method | `apps/powerpoint-vsto/UniversalConnectors.cs:114` |
| `Selected` | method | `apps/powerpoint-vsto/UniversalConnectors.cs:118` |
| `Update` | method | `apps/powerpoint-vsto/UniversalConnectors.cs:119` |
| `Reflow` | method | `apps/powerpoint-vsto/UniversalConnectors.cs:123` |
| `Insert` | method | `apps/powerpoint-vsto/UniversalConnectors.cs:132` |
| `Edit` | method | `apps/powerpoint-vsto/UniversalConnectors.cs:133` |
| `EndpointEditor` | class | `apps/powerpoint-vsto/UniversalConnectors.cs:134` |
| `Read` | method | `apps/powerpoint-vsto/UniversalConnectors.cs:144` |
| `Show` | method | `apps/powerpoint-vsto/UniversalConnectors.cs:146` |
