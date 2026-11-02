# Ribbon deployment regression investigation — 22 October 2026

Goal remains active: test all new features and verify their installed ribbon UI.

Observed package before changes: bin/x64/Release/PowerPointIntelligence.dll was
54,272 bytes dated 11:55:38; VSTO deployment version remained 1.0.0.0 with manifests
dated 12:06. Those artifacts preceded the demo source. Latest installer transcript
contained only the installation step, not a new build. Compile targets update
intermediate assemblies; they do not establish a fresh signed deployment.

Changes: compiled DemoRibbon generator now exposes all 90 local command IDs in
categorized menus; actual generated XML is tested rather than reconstructing a
C# string by regex. Every catalog ID has exactly one RunDemoTool binding and valid
callback, unique ID and namespace. All 103 inventory identities are checked.

Installer now rejects missing current demo marker, source newer than the DLL,
and SHA-256 mismatches between DLL/application/deployment manifests. The actual
old package fails with STALE PACKAGE. On successful installation, a receipt records
assembly hash, manifest path, deployment version and installer exit code. No
registry or cache is forcibly cleared and no security policy is bypassed.

Installed runtime check verifies connected COM add-in, loaded marker/hash against
receipt, actual ribbon OnLoad callback and 90 generated bindings through the
add-in's diagnostic API. It is not a screenshot or visual layout assertion.
Attempt in this session fails at PowerPoint activation, 0x80070005.

A per-feature Office harness is compiled and integrated in Test-All. It enumerates
all commands and records PASS with scoped assertions, SMOKE_ONLY, MANUAL_REQUIRED,
FAIL or BLOCKED. Dialog/file interactions are not claimed tested. Attempt here
records all 90 BLOCKED due to PowerPoint activation denial, rather than counting
skipped features as passes. User-requested all-feature runtime acceptance is
therefore incomplete. The harness preserves user presentations and does not quit
PowerPoint. New tests need desktop execution with a freshly rebuilt installation.
