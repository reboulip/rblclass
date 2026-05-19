# CLAUDE.md

## Project overview

This is a refactor of a legacy Outlook VBA macro into a modern VSTO add-in.
It helps users quickly classify emails into folders across multiple .pst
archives, with fast keyword search over the folder tree, attachment
management, and send-time guards.

## Stack and constraints

- **Add-in technology**: VSTO (Visual Studio Tools for Office). This is
  non-negotiable: it is the only technology supporting PST access, offline
  use, and per-user install without admin rights on the classic Win32
  Outlook client.
- **Target Outlook**: classic Win32 Outlook, **32-bit**, Microsoft 365
  Apps for Enterprise on the **Semi-Annual Enterprise Channel**.
  NOT the "new Outlook". NO Exchange Online dependency.
- **Runtime**: .NET Framework 4.8 for the VSTO add-in and the Outlook
  adapter. .NET Standard 2.0 for the business core.
- **Language**: C# 7.3 (.NET Framework 4.8 limit unless LangVersion is
  explicitly bumped, which we avoid).
- **UI**: WPF with MVVM via CommunityToolkit.Mvvm. Ribbon via Ribbon XML
  (not the Ribbon Designer).
- **Storage**: SQLite via Microsoft.Data.Sqlite, with FTS5 for full-text
  search. Database in %LocalAppData%\RBLclass\.
- **Logging**: Serilog with a rolling file sink.
- **Tests**: xUnit, FluentAssertions, NSubstitute. Tests run against
  RBLclass.Core only; the Outlook adapter is excluded from CI tests.
- **Packaging**: ClickOnce, per-user install, code-signed with an
  internal-PKI certificate carrying the Code Signing EKU
  (1.3.6.1.5.5.7.3.3).
- **Min Windows**: Windows 10 1809.

## Architecture

Strict layering. Dependencies point downward only.

RBLclass.VstoAddin           (.NET FW 4.8) — Ribbon, Task Pane, events
       │
RBLclass.Outlook.Adapter     (.NET FW 4.8) — COM access to Outlook OM
       │                                  Implements RBLclass.Core interfaces
       ▼
RBLclass.Core                (.NET Standard 2.0) — business logic, no Outlook
                                                no UI, no I/O except via
                                                injected interfaces


The business core (`RBLclass.Core`) MUST remain free of any reference to
Microsoft.Office.Interop.Outlook, System.Windows, or VSTO assemblies.
This is what makes the code portable on the day we have to migrate away
from VSTO.

## Critical coding rules

### COM lifetime (Outlook adapter only)

Outlook COM objects MUST be released explicitly. Long-lived references
cause Outlook crashes and memory leaks.

- Wrap every COM object in a `ComRef<T> : IDisposable` that calls
  `Marshal.ReleaseComObject` in `Dispose`.
- Never chain property accesses on COM objects. `mail.Parent.Parent.Name`
  leaks the intermediate `Parent` objects.
- All loops over `Items`, `Folders`, `Stores` use `for` with index, not
  `foreach`, and dispose each element.

### Threading

- Outlook OM is single-threaded apartment (STA). All COM calls happen on
  the main thread.
- Background work (indexing, SQLite I/O, HTTP) runs on the thread pool.
- Use a dedicated `SynchronizationContext` capture at add-in startup to
  marshal back to the Outlook UI thread.
- Never block the UI thread on COM calls inside loops. Batch and yield.

### Bitness

- Outlook runs as a **32-bit process** on target workstations. The VSTO
  add-in therefore loads into a 32-bit host.
- Build configuration: `x86` for `RBLclass.VstoAddin` and
  `RBLclass.Outlook.Adapter`. `AnyCPU` is acceptable for `RBLclass.Core` and
  test projects.
- Any native dependency (SQLite, etc.) MUST ship with its **x86 native
  binary**. Use the NuGet packages that include both x86 and x64
  payloads (e.g. `SQLitePCLRaw.bundle_e_sqlite3`) and verify after build
  that `e_sqlite3.dll` (x86) is present in the output directory.
- Do NOT add dependencies that are x64-only.
- Memory budget: 32-bit processes are capped at ~2 GB address space
  shared with Outlook itself. Avoid loading large datasets in memory;
  stream from SQLite.

### Performance targets

- Folder search: <100ms for a query on 100k indexed mails.
- Classify action: optimistic UI, perceived latency <300ms. Actual Move
  may complete asynchronously.
- Add-in startup: contribute <500ms to Outlook startup. Initial PST
  indexing runs in background after Outlook is fully loaded.

### SQLite

- One connection per logical operation, opened with WAL journal mode.
- All schema changes go through versioned migrations (`SchemaVersion`
  table). Never alter the schema in-place.
- FTS5 virtual table for mail body/subject. Keep it separate from the
  main `Mails` table to allow rebuilds without data loss.
- EntryIDs from Outlook can change after some operations (e.g. moving
  between stores). Cache by `(StoreID, EntryID)` and tolerate misses.

### Outlook events

- Register event handlers in `ThisAddIn_Startup`, unregister in
  `ThisAddIn_Shutdown`.
- Keep references to event source objects in fields, otherwise GC will
  collect them and events stop firing silently.

### Error handling

- Outlook COM calls can throw `COMException`. Catch and log, never let
  exceptions escape to Outlook (it disables the add-in).
- Top-level handlers in every ribbon callback, every event handler, every
  fire-and-forget Task.

## Folder layout

```
/src
  /RBLclass.Core              business logic, .NET Standard 2.0
  /RBLclass.Outlook.Adapter   COM adapter, .NET Framework 4.8
  /RBLclass.VstoAddin         VSTO shell, .NET Framework 4.8
/tests
  /RBLclass.Core.Tests        xUnit
/docs
  /architecture.md
  /derisking.md
ROADMAP.md
CLAUDE.md
README.md
```

## What Claude should do

- When asked to add a feature, locate it in the appropriate layer. If it
  needs Outlook data, define an interface in `RBLclass.Core`, implement it
  in `RBLclass.Outlook.Adapter`, write the business logic in `RBLclass.Core`,
  wire it from `RBLclass.VstoAddin`.
- Always write unit tests for any non-trivial logic added in `RBLclass.Core`.
- Follow existing naming, formatting (`.editorconfig`), and namespace
  conventions.
- Prefer small, focused changes over large refactors. If a refactor is
  needed, propose it first, do not perform it unprompted.

## What Claude should NOT do

- Do not add references from `RBLclass.Core` to anything Outlook-, WPF-, or
  VSTO-related.
- Do not introduce new dependencies without asking. Keep the dependency
  surface small.
- Do not call Outlook COM APIs from a non-UI thread.
- Do not use `async void` except for event handlers, and even then wrap
  the body in try/catch.
- Do not bump LangVersion or TargetFramework without an explicit request.
- Do not assume Exchange Online is available. PST is the source of truth.
- Do not introduce x64-only native dependencies.
- Do not assume more than ~1 GB of usable memory; the add-in lives
  inside a 32-bit Outlook process.

## Useful context for Outlook OM quirks

- `MailItem.Move` returns the moved item in the destination store; the
  original COM reference becomes invalid.
- `Explorer.Selection` can contain non-MailItem objects (MeetingItem,
  ReportItem...). Always check `is MailItem` before casting.
- `Folder.Items.Find` / `Restrict` is much faster than iterating, but
  has its own DASL/Jet query syntax — document any non-trivial query.
- PST stores can be closed by the user at any time. Handle
  `StoreRemove` events gracefully.

## Build and run

- Open `RBLclass.sln` in Visual Studio 2022 with the "Office/SharePoint
  development" workload installed.
- Set `RBLclass.VstoAddin` as startup project. F5 launches Outlook with the
  add-in attached.
- `RBLclass.Core` and its tests can be built and run from VS Code with the
  C# Dev Kit; do not try to build the VSTO project from VS Code.
- Default build configuration is `Debug|x86`. Release is `Release|x86`.
  The `AnyCPU` configuration exists only for `RBLclass.Core` and test
  projects.

## Publishing

- ClickOnce profile in `RBLclass.VstoAddin\Properties\PublishProfiles`.
- Manifest is signed with the certificate referenced in
  `RBLclass.VstoAddin.csproj` (`<ManifestCertificateThumbprint>`).
- Publish location: internal HTTPS share documented in
  `docs/deployment.md`.

## Repository management

Always work on develop branch or a branch based on develop.