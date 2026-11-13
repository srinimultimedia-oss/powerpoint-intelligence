# Shared selected-input and generated-output policy

**VERIFIED source review.** This design describes inspected code, not current runtime/Office execution.

## 1. Purpose

Separates selected external inputs, independently approved roots and generated workspace paths.

## 2. Implementation status

Production linked utility and test outputguard; Windows semantics.

## 3. Source locations

- `src/Library.Core/PathPolicy.cs:1`

## 4. Internal components

System.IO/Reflection; no network calls. Reject dot segments, controls, trailing spaces, device names, ADS and wildcard syntax; allow fully qualified drive/UNC selected inputs; reconstruct child components under root with separator-bounded comparison. Generated configuration, index and test paths reject existing linked ancestors. TestPath discovers repository marker/OS temp and forbids the workspace root.

## 5. Classes and interfaces

- `src/Library.Core/PathPolicy.cs:12` — `internal static class PathPolicy`

## 6. Important methods and functions

Declaration locator below is lexical and may include overload/call-like lines; semantic workflow is described separately.

- `src/Library.Core/PathPolicy.cs:35` — `internal static string Selected(string path)`
- `src/Library.Core/PathPolicy.cs:48` — `internal static bool IsWithin(string root, string candidate)`
- `src/Library.Core/PathPolicy.cs:56` — `internal static string Within(string root, string path)`
- `src/Library.Core/PathPolicy.cs:76` — `internal static string ApprovedRoot(string requested, IEnumerable<string> approvedRoots)`
- `src/Library.Core/PathPolicy.cs:89` — `internal static string Configured(string value, string baseDirectory)`
- `src/Library.Core/PathPolicy.cs:94` — `internal static void RejectLinkedAncestors(string path)`
- `src/Library.Core/PathPolicy.cs:124` — `internal static string TestPath(string path)`
- `src/Library.Core/PathPolicy.cs:141` — `internal static string ConfigFile(string path)`

## 7. Inputs and outputs

Path plus trusted base/roots -> reconstructed fullpath or InvalidDataException.

## 8. Dependencies

System.IO/Reflection; no network calls.

## 9. Workflow

Reject dot segments, controls, trailing spaces, device names, ADS and wildcard syntax; allow fully qualified drive/UNC selected inputs; reconstruct child components under root with separator-bounded comparison. Generated configuration, index and test paths reject existing linked ancestors. TestPath discovers repository marker/OS temp and forbids the workspace root.

## 10. Integration points

Path plus trusted base/roots -> reconstructed fullpath or InvalidDataException. Production linked utility and test outputguard; Windows semantics.

## 11. Exceptions and recovery

InvalidDataException before filesystem sinks on unsupported path; caller preserves original objects.

## 12. Configuration dependencies

32,767-character path cap; configuration uses .config/.xml only.

## 13. Security considerations

Legacy requested root is compared against independent registrations; selected files maylive outside the repository by design, generated outputs cannot.

## 14. Unit and integration test definitions

- `tests/security/PathPolicyTests.cs:1`
- `tests/s3/SceneArtifactPath.cs:1`

Inspected definitions only. See the parent current test report for execution; historical passes do not establish current acceptance.

## 15. Known limitations

Name/attribute validation is not a complete handle-based race-proof sandbox; selected input link policy intentionally differs.

## 16. Technical debt

Boundary changes affect many tools; device/UNC/longpath negative fixtures important.

## 17. Extension points

Add path forms onlywith positive/negative authority tests and outside-output preservation.
