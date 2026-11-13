# Application architecture

## Startup and lifecycle

VSTO creates the generated `ThisAddIn` host item and calls the application's startup handler. That handler constructs the host adapter, starts layout timing, prepares an agenda timer, evaluates Office compatibility, creates the context-property controller, subscribes selection/double-click/open/save/activate/close events and invalidates an existing ribbon. The 750 ms layout and 650 ms agenda intervals are WinForms UI scheduling, not backend workers. `apps/powerpoint-vsto/ThisAddIn.cs:37`.

Compatibility loads the deployed DLL configuration and compares API major 16, requested architecture and optional executable build floor. An Office API string of 16.0 cannot distinguish 2016, 2019 and Microsoft 365 licensing/update channels. Invalid configuration disables document commands rather than guessing. `apps/powerpoint-vsto/ThisAddIn.cs:52`; `apps/powerpoint-vsto/OfficeCompatibility.cs:1`.

## Command execution

Ribbon callback objects forward to `ThisAddIn`, which applies `busy`, host, compatibility and mutation-lease checks. `EnsurePane` creates/rebinds a task pane to the current window on demand. User action code executes while holding a named mutex for the PowerPoint PID. Specific failures become status text; a generic catch reports operation unavailability. Cancellation is a no-op. The `finally` path resets busy and requests a context refresh. `apps/powerpoint-vsto/ThisAddIn.cs:133`.

Selection changes request a property snapshot, invalidate controls and restart the debounce timer. Activation avoids creating/discarding panes or rewriting metadata directly in the COM callback; it marks pane refresh. Before-save work targets only the active presentation and refreshes managed agenda/layout/chart state. These handlers use an exception boundary to prevent managed exceptions escaping Office callbacks. `apps/powerpoint-vsto/ThisAddIn.cs:318`; `apps/powerpoint-vsto/PowerPointHost.cs:34`.

## Shutdown and resource ownership

Shutdown disposes the context controller, timers and layout form, unsubscribes all registered application events and removes the task pane. Native COM object lifetime is handled locally in modules; the repository does not implement a general reference-count tracker proving every RCW release. Never release shared Office application objects indiscriminately. `apps/powerpoint-vsto/ThisAddIn.cs:77`. See [interoperability](../04-powerpoint-integration/interoperability.md).

## Runtime limits

The floating property box is foreground/selection-sensitive and can hide while Excel owns focus. Mutation coordination does not make COM calls thread-safe. Live service operations, enterprise scheduling, authorization and model invocation are not implemented merely because command labels exist. Details are mapped in [feature status](../05-features/feature-status-matrix.md).
