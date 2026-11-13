# Implemented AI-adjacent workflow

Only profile preparation is implemented. The following sequence is directly supported by `apps/powerpoint-vsto/ThisAddIn.cs:252` and `apps/powerpoint-vsto/ServiceConnections.cs:98`.

```mermaid
sequenceDiagram
  actor User
  participant Ribbon
  participant AddIn as ThisAddIn
  participant Profiles as ServiceConnections
  participant Disk as Local XML file
  User->>Ribbon: Configure services
  Ribbon->>AddIn: RunWorkflow(connections)
  AddIn->>Profiles: ShowConfiguration()
  Profiles->>Disk: Load(default path)
  alt Missing configuration
    Profiles->>Profiles: Three empty profiles
  else Invalid or inaccessible configuration
    Profiles->>Profiles: Empty profiles plus warning
  end
  User->>Profiles: Edit nonsecret metadata
  Profiles->>Profiles: Validate and show readiness label
  alt Cancel
    Profiles-->>User: Existing file preserved
  else Save valid metadata
    Profiles->>Disk: Write temporary XML then replace/move
    Profiles-->>User: Close settings dialog
  else Invalid or denied save
    Profiles-->>User: Error; no success response
  end
```

The sequence deliberately contains no Azure/Entra/network actor. A malformed existing configuration is not silently repaired on open: Cancel preserves it; an explicit valid Save replaces it. Temporary save files are removed in a finally block. `apps/powerpoint-vsto/ServiceConnections.cs:93`.

The README's request → RBAC → retrieval → Azure → plan validation → approval → host → audit chain remains **PLANNED**. No live AI request-processing diagram can be derived from implemented source.
