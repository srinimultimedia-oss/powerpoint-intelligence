# NexWave display-name update — 3 November 2026

Owner request: rename the plugin from Intelligence to **NexWave**.

Updated the ribbon tab and keyboard keytip (NW), task panes, library/gallery titles, contextual properties, agenda/chart/series/connector/text dialogs, newly created selection-pane shape names, applicable messages/filter labels, VSTO FriendlyName/ProductName, installer prompts and current/generated README headings. A new build marker identifies the renamed payload. Automation test selectors retain the stable ProgID match and recognize the new description.

Compatibility identities are unchanged: project/assembly/namespace, ProgID/project GUID, persisted XML namespaces and shape tags, per-user data/index folders, embedded payload/resource identity, setup EXE/VSTO filenames and build mutex. Existing signing-certificate subject and publisher trust are unchanged; no new certificate or trust-store operation was performed. Existing presentations and historical evidence were not rewritten.

The source now builds the display name NexWave; the running installed add-in and previously signed EXE are not automatically renamed. This session has no accessible private code-signing certificate, so no new signed installer or native PowerPoint loading is claimed. Save and close presentations, run scripts/Install-Local.ps1 from normal desktop PowerShell, then reopen PowerPoint to verify the NexWave tab. Rebuild easySetup through its normal approved signing workflow if distributing the EXE.

Verification: compiled ribbon XML shows label NexWave and keytip NW; VSTO project XML FriendlyName/ProductName both NexWave while AssemblyName stays PowerPointIntelligence. Ribbon/tool/icon contracts pass. Template library tests and setup SourceOnly checks (8 policy + 3 resource-reader assertions) pass. Full final local evidence is recorded below. Native Office display remains a desktop acceptance gate.

Final validation: `artifacts/validation-20261103-092949-92a2cd84/results.json`, source `8063EE6F962C09A84AB4093408EF0DB4F1C6159F14B57790978A328EEA726891`, matching the current source. **23 PASS / 1 FAIL**; the sole failure is the already-missing PIIDX test DLL protected by its recovery gate. All three architecture compile/compatibility targets, ribbon/command contracts, renderer/scene/security and template tests pass. See `artifacts/nexwave-ribbon.log`, `nexwave-template.log` and `nexwave-setup-source.log`. No missing binary was restored or native Office acceptance claimed.
