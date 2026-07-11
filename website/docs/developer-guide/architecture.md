# Architecture

RBLclass is **strictly layered**. Dependencies point downward only.

```text
RBLclass.AddIn              (.NET FW 4.8)  COM add-in shell:
      │                                    IDTExtensibility2, IRibbonExtensibility,
      │                                    ribbon callbacks, Custom Task Panes,
      │                                    Outlook event subscriptions
RBLclass.Outlook.Adapter    (.NET FW 4.8)  COM access to the Outlook Object Model,
      │                                    implements RBLclass.Core interfaces
      ▼
RBLclass.Core               (.NET Standard 2.0)  business logic — no Outlook,
                                             no UI, no I/O except via injected
                                             interfaces
```

!!! abstract "The one rule that matters most"
    `RBLclass.Core` must **never** reference
    `Microsoft.Office.Interop.Outlook`, `System.Windows`, or the COM-add-in shell
    assemblies. That is what keeps the business logic portable on the day the COM
    add-in model has to be left behind.

When a feature needs Outlook data, the flow is always the same:

1. Define an **interface** in `RBLclass.Core`.
2. Implement it in `RBLclass.Outlook.Adapter` (the COM boundary).
3. Write the business logic in `RBLclass.Core` against the interface.
4. Wire it up from `RBLclass.AddIn`.

## Folder layout

```text
/src
  /RBLclass.Core              business logic, .NET Standard 2.0
  /RBLclass.Outlook.Adapter   COM adapter, .NET Framework 4.8
  /RBLclass.AddIn             COM add-in shell, .NET Framework 4.8
/installer                    WiX MSI project (Phase 1)
/tests
  /RBLclass.Core.Tests        xUnit
/website                      this documentation site (MkDocs Material)
/docs                         architecture notes, install guide, roadmap detail
ROADMAP.md  CLAUDE.md  README.md
```

Files most often relevant when changing behaviour:

- `src/RBLclass.AddIn/RblClassAddIn.cs` — COM entry points, ribbon callbacks, CTP
  management, event registration, startup.
- `src/RBLclass.AddIn/ViewModels/MainPaneViewModel.cs` — main pane logic.
- `src/RBLclass.AddIn/Views/MainPaneView.xaml` (+ `.xaml.cs`) — pane XAML.
- `src/RBLclass.Core/Settings.cs` + `SettingsKeys.cs` — typed settings.
- `src/RBLclass.AddIn/Ribbon.xml` (+ `.fr.xml`, `.de.xml`) — ribbon definition.

## COM interop rules (shell only)

These conventions exist because getting them wrong takes Outlook down silently
and is very hard to diagnose:

- **Never hand-roll `[ComImport]` declarations** of `IDTExtensibility2`,
  `IRibbonExtensibility`, `IRibbonControl`, etc. Reference the canonical PIAs
  from the GAC (`Extensibility`, `office`,
  `Microsoft.Office.Interop.Outlook`) with `<Private>false</Private>`. The CLR
  auto-generates marshalling stubs from interface metadata; subtle deviations
  from the canonical PIAs produce stubs that access-violate inside
  `IL_STUB_COMtoCLR`, surfacing as `ExecutionEngineException` (HRESULT
  `0x80131506`).
- **`[ClassInterface(ClassInterfaceType.AutoDispatch)]` on the add-in class is
  mandatory** — not `None`. Office resolves ribbon `onAction` callbacks via
  `IDispatch::GetIDsOfNames` on the class itself.
- **Do not embed Office interop types** — use direct GAC references, which are
  never copied to the output.

## COM lifetime (adapter only)

Outlook COM objects must be released explicitly, or Outlook leaks and crashes:

- Wrap every COM object in a `ComRef<T> : IDisposable` that calls
  `Marshal.ReleaseComObject` in `Dispose`.
- **Never chain property accesses** on COM objects — `mail.Parent.Parent.Name`
  leaks the intermediate `Parent` objects.
- Loop over `Items` / `Folders` / `Stores` with a `for` index (not `foreach`),
  disposing each element.

## Threading

- The Outlook OM is single-threaded apartment (STA); **all COM calls happen on
  the main thread**.
- Background work (indexing, SQLite I/O, HTTP) runs on the thread pool.
- A `SynchronizationContext` captured at add-in startup marshals results back to
  the Outlook UI thread.
- Never block the UI thread on COM calls inside loops — batch and yield.

## Storage & SQLite policy

- One connection per logical operation, opened with **WAL** journal mode.
- All schema changes go through **versioned migrations** (`SchemaVersion`
  table) — never alter the schema in-place.
- The **FTS5** virtual table for search is kept separate from the main `Mails`
  table so it can be rebuilt without data loss.
- EntryIDs can change after some Outlook operations (e.g. moving between stores).
  Cache by `(StoreID, EntryID)` and **tolerate misses**.

## Outlook events

- Register handlers in `OnStartupComplete`, unregister in `OnBeginShutdown`.
- Keep references to event source objects in **fields** — otherwise GC collects
  them and events stop firing silently.
- Handle `StoreRemove` gracefully: PST stores can be closed by the user at any
  time.

## Error handling

- Outlook COM calls throw `COMException`. **Catch and log; never let an exception
  escape into Outlook** — it disables the add-in by flipping `LoadBehavior` from
  3 to 2 in the registry.
- Top-level handlers wrap every ribbon callback, every event handler, and every
  fire-and-forget `Task`.

## Localization

- XAML strings go through `{loc:Loc Key=...}`
  (`RBLclass.AddIn.Localization.LocExtension`).
- Code-behind / ViewModel strings go through
  `ILocalizationService.GetString` / `GetString(key, args)`.
- Count-dependent strings use a `_One` / `_Other` key pair resolved via
  `ILocalizationService.Plural(...)` — never concatenate, because French and
  German plural forms differ from English.
- Ribbon buttons/tooltips need matching entries in all three `Ribbon*.xml`
  files. `ResourceParityTests` enforces parity across `.resx` and ribbon files.

## Performance targets

| Operation | Target |
| --- | --- |
| Folder search | < 100 ms for a query over 100k indexed mails |
| Classify action | Optimistic UI, perceived latency < 300 ms (the actual Move may complete asynchronously) |
| Add-in startup | Contributes < 500 ms to Outlook startup; initial PST indexing runs in the background afterwards |

## Outlook OM quirks worth remembering

- `MailItem.Move` returns the moved item **in the destination store**; the
  original COM reference becomes invalid.
- `Explorer.Selection` can contain non-`MailItem` objects (`MeetingItem`,
  `ReportItem`, …) — always check `is MailItem` before casting.
- `Folder.Items.Find` / `Restrict` is much faster than iterating, but uses its
  own DASL/Jet query syntax — document any non-trivial query.

The full, authoritative version of these rules lives in
[`CLAUDE.md`](https://github.com/reboulip/rblclass/blob/main/CLAUDE.md) at the
repo root.
