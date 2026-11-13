# Request/response and file contracts

## S0 COM JSON request (implemented)

```json
{"SchemaVersion":1,"ChartId":"11111111-1111-4111-8111-111111111111","ChartType":"clustered-column","Revision":0,"Categories":["Category A"],"Series":[{"Id":"s1","Name":"Series A","Values":[12.5]}]}
```

This sanitized example uses real ChartDocument/ChartSeries property names. InsertChartJson validates through MetadataCodec, then replaces supplied ChartId with a new GUID and resets Revision to zero before insertion. Return is the new chart-ID string, not an HTTP JSON envelope. Nullable numeric Values represent missing values; dimension/type/identity bounds are enforced by MetadataCodec. `apps/powerpoint-vsto/Contracts.cs:17` `apps/powerpoint-vsto/DemoApi.cs:41`.

## Service profile XML (implemented)

Root is Connections in `urn:powerpoint-intelligence:service-connections:v1` with version=1. It has exactly three Profile children; attributes service/authentication and text-only Endpoint/Deployment children. Empty endpoints are permitted for unconfigured services. A profile cannot contain an ApiKey/password field. Exact codec and constraints: `apps/powerpoint-vsto/ServiceConnections.cs:58` `apps/powerpoint-vsto/ServiceConnections.cs:75`.

## AI action plan (schema only)

```json
{"planId":"example-plan","status":"proposed","approvalRequired":true,"actions":[{"actionId":"action-1","type":"CREATE_CHART","arguments":{}}]}
```

This is valid against the **planned** schema shape; it is not a successful AI response or executable chart request. Arguments has no defined chart arguments yet. There is no route, authentication, provider payload, retry or response parser attached. `services/ai-gateway/action-plan.schema.json:1`.

## Other versioned boundaries

| Contract | Authority | Design reference |
|---|---|---|
| Canonical chart v2/v3 | Chart.Domain codecs and JSON schema | [Serialization contracts](../08-data-design/serialization-contracts.md) |
| Rendered scene v1 | SceneCodec/StrictSceneJson/BoundedJson; semantic validation against canonical model | [Data models](../08-data-design/data-models.md) |
| Native rendered metadata | RenderedChartMetadata with canonical JSON, scene JSON and bindings | [Persistence](../08-data-design/persistence-design.md) |
| Linked tables/source links/mappings | DataSources, ExcelRangeMapping and DataLinkManager | [Feature details](../05-features/feature-implementation-details.md) |
| PIIDX catalog | SQLite tables/metadata/FTS and source extraction | [Data dictionary](../08-data-design/data-dictionary.md) |
| Installer receipt | build.json containing version, source/EXE/bundle SHA256, signer thumbprint, architectures/template count/time | [Packaging](../11-build-deployment/packaging-installation.md) |

Models are not interchangeable solely because they contain IDs/data. Use the exact codec and migration path at each boundary; native S0 automation does not automatically accept v2/v3 JSON.
