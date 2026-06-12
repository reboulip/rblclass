# CLAUDE.md

## Project overview

This is a refactor of a legacy Outlook VBA macro into a modern Outlook
COM add-in. It helps users quickly classify emails into folders across
multiple .pst archives, with fast keyword search over the folder tree,
attachment management, and send-time guards.

## Stack and constraints

- **Add-in technology**: Outlook **Shared COM Add-in** —
  `Extensibility.IDTExtensibility2` + `Office.IRibbonExtensibility`,
  registered into Outlook via HKCU keys under
  `Software\Microsoft\Office\Outlook\Addins\`. This is the classic,
  pre-VSTO Outlook add-in model. It supports PST access, offline use,
  and per-user install without admin rights on the classic Win32
  Outlook client. VSTO is explicitly excluded for this project after
  the Office Developer Tools workload could not be set up on the dev
  machine.
- **Target Outlook** (deployment): classic Win32 Outlook, **32-bit**,
  Microsoft 365 Apps for Enterprise on the **Semi-Annual Enterprise
  Channel**. NOT the "new Outlook". NO Exchange Online dependency.
- **Dev Outlook**: classic Win32 Outlook, **64-bit**, **Current**
  channel, on **Windows 11 ARM64** (Outlook runs as emulated x64
  via Prism). The local dev machine differs from the deployment
  target in both architecture and channel. The same AnyCPU
  add-in binary must load on both. Final EDR and compatibility
  validation happens against a real 32-bit target workstation;
  the dev machine is for fast iteration only.
- **Runtime**: .NET Framework 4.8 for the add-in shell and the
  Outlook adapter. .NET Standard 2.0 for the business core.
- **Language**: C# 7.3 (.NET Framework 4.8 limit unless LangVersion is
  explicitly bumped, which we avoid).
- **UI**: WPF with a **hand-rolled minimal MVVM** (`RBLclass.AddIn/Mvvm`:
  `ObservableObject` + `RelayCommand`), hosted in Custom Task Panes via
  `ICustomTaskPaneConsumer`/`ICTPFactory` — the WPF view is bridged into a
  ComVisible WinForms host control through `ElementHost`. Ribbon via
  `IRibbonExtensibility.GetCustomUI` returning Ribbon XML (not the
  Ribbon Designer). **CommunityToolkit.Mvvm is deliberately NOT used:** its
  8.x dependencies (`System.Memory` ≥ 4.5.5, `System.Runtime.CompilerServices.Unsafe`
  ≥ 6.0.0) float the SQLite facade assemblies above what `SQLitePCLRaw` was
  built against, and a COM host has no binding redirects, so it breaks the
  first `SqliteConnection` at runtime. Keep new AddIn dependencies
  dependency-light for the same reason.
- **Storage**: SQLite via Microsoft.Data.Sqlite, with FTS5 for full-text
  search. Database in %LocalAppData%\RBLclass\.
- **Logging**: Serilog with a rolling file sink.
- **Tests**: xUnit, FluentAssertions, NSubstitute. Tests run against
  RBLclass.Core only; the Outlook adapter is excluded from CI tests.
- **Packaging**: per-user install via a signed installer that lays
  down the add-in DLL + dependencies under `%LocalAppData%\RBLclass\`
  and writes the COM registration entries to HKCU. Phase 0 POC uses a
  PowerShell installer; Phase 1 promotes to an MSI built with the WiX
  Toolset, Authenticode-signed (both the DLL and the MSI) with an
  internal-PKI certificate carrying the Code Signing EKU
  (1.3.6.1.5.5.7.3.3). ClickOnce is not used — it is VSTO-specific
  in this scenario.
- **Min Windows**: Windows 10 1809.

## Architecture

Strict layering. Dependencies point downward only.

RBLclass.AddIn               (.NET FW 4.8) — COM add-in shell:
       │                                  IDTExtensibility2, IRibbonExtensibility,
       │                                  ribbon callbacks, Custom Task Panes,
       │                                  Outlook event subscriptions
RBLclass.Outlook.Adapter     (.NET FW 4.8) — COM access to Outlook OM
       │                                  Implements RBLclass.Core interfaces
       ▼
RBLclass.Core                (.NET Standard 2.0) — business logic, no Outlook
                                                no UI, no I/O except via
                                                injected interfaces


The business core (`RBLclass.Core`) MUST remain free of any reference to
Microsoft.Office.Interop.Outlook, System.Windows, or COM-add-in shell
assemblies. This is what makes the code portable on the day we have to
migrate away from the COM add-in model.

## Critical coding rules

### COM interop interface declarations (add-in shell only)

- **Never hand-roll `[ComImport]` declarations of `IDTExtensibility2`,
  `IRibbonExtensibility`, `IRibbonControl`, or other Office/Extensibility
  interfaces.** Reference the canonical PIAs from the GAC instead:
  - `Extensibility` —
    `C:\Windows\assembly\GAC\Extensibility\7.0.3300.0__b03f5f7f11d50a3a\extensibility.dll`
  - `office` (Microsoft.Office.Core) —
    `C:\Windows\assembly\GAC_MSIL\office\15.0.0.0__71e9bce111e9429c\OFFICE.DLL`
  - `Microsoft.Office.Interop.Outlook` —
    `C:\Windows\assembly\GAC_MSIL\Microsoft.Office.Interop.Outlook\15.0.0.0__71e9bce111e9429c\Microsoft.Office.Interop.Outlook.dll`

  Reference them in the `.csproj` via `<Reference Include="…"><HintPath>…</HintPath><Private>false</Private></Reference>`.
  **Why:** CLR auto-generates IL marshalling stubs from interface
  metadata. Even with the correct `[Guid]` and `[DispId]` attributes,
  subtle differences from the canonical PIAs (e.g. missing `[In]`,
  wrong `MarshalAs` for `ref Array`) produce stubs that AV inside
  `IL_STUB_COMtoCLR` when Outlook calls our methods. The crash
  surfaces as `ExecutionEngineException` (HRESULT `0x80131506`),
  takes Outlook down silently, and is hard to diagnose without
  a memory dump.

- **`[ClassInterface(ClassInterfaceType.AutoDispatch)]` on the add-in
  class is mandatory** — not `None`. Office resolves ribbon
  `onAction` callbacks by calling `IDispatch::GetIDsOfNames` on
  the class itself, and `None` only exposes members of explicitly-
  implemented interfaces, none of which the ribbon callbacks live on.

- **Do not embed Office interop types** (`<EmbedInteropTypes>true</EmbedInteropTypes>`
  on a `PackageReference`). In SDK-style net48 projects this metadata
  is silently ignored and the PIA is copied as a regular dependency.
  Use direct GAC references (as above) which never get copied.

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

- Build configuration: **AnyCPU** for all projects, with
  `Prefer32Bit=false`. The same binary loads into 64-bit Outlook on
  the ARM64 dev machine (running as x64 under Prism emulation) AND
  32-bit Outlook on the x86/x64 target workstations; the .NET
  runtime JITs to the host process bitness.
- Native dependencies MUST ship both `runtimes\win-x86\native\` AND
  `runtimes\win-x64\native\` payloads (e.g.
  `SQLitePCLRaw.bundle_e_sqlite3`). Verify after build that both
  payloads land in the output directory. Do NOT add x64-only
  dependencies.
- COM registration is bitness-specific. Under HKCU on 64-bit
  Windows, 64-bit clients read `Software\Classes\CLSID\{guid}` while
  32-bit clients read `Software\Classes\Wow6432Node\CLSID\{guid}`.
  The installer writes both subtrees so a single artifact supports
  both Outlook bitnesses without per-host customisation. The
  `Software\Microsoft\Office\Outlook\Addins\` key is NOT WOW64-
  redirected in HKCU and is written once.
- Memory budget: 32-bit processes are capped at ~2 GB address space
  shared with Outlook itself. Avoid loading large datasets in memory;
  stream from SQLite. Treat ~1 GB as the usable budget.

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

- Register event handlers in `OnStartupComplete`, unregister in
  `OnBeginShutdown`.
- Keep references to event source objects in fields, otherwise GC will
  collect them and events stop firing silently.

### Error handling

- Outlook COM calls can throw `COMException`. Catch and log, never let
  exceptions escape to Outlook (it disables the add-in by flipping
  `LoadBehavior` from 3 to 2 in the registry).
- Top-level handlers in every ribbon callback, every event handler, every
  fire-and-forget Task.

## Folder layout

```
/src
  /RBLclass.Core              business logic, .NET Standard 2.0
  /RBLclass.Outlook.Adapter   COM adapter, .NET Framework 4.8
  /RBLclass.AddIn             COM add-in shell, .NET Framework 4.8
/installer                    WiX MSI project (Phase 1)
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
  wire it from `RBLclass.AddIn`.
- Always write unit tests for any non-trivial logic added in `RBLclass.Core`.
- Follow existing naming, formatting (`.editorconfig`), and namespace
  conventions.
- Prefer small, focused changes over large refactors. If a refactor is
  needed, propose it first, do not perform it unprompted.

## What Claude should NOT do

- Do not add references from `RBLclass.Core` to anything Outlook-, WPF-, or
  COM-add-in-shell-related.
- Do not depend on the VSTO runtime or the `Microsoft.Office.Tools.*`
  assemblies — they are not used in this project.
- Do not introduce new dependencies without asking. Keep the dependency
  surface small.
- Do not call Outlook COM APIs from a non-UI thread.
- Do not use `async void` except for event handlers, and even then wrap
  the body in try/catch.
- Do not bump LangVersion or TargetFramework without an explicit request.
- Do not assume Exchange Online is available. PST is the source of truth.
- Do not introduce x64-only native dependencies.
- Do not assume more than ~1 GB of usable memory; the add-in lives
  inside a 32-bit Outlook process on the deployment target.

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

- VS 2022 with the **.NET desktop development** workload and the
  **.NET Framework 4.8 targeting pack** is required. The
  Office/SharePoint development workload is NOT required.
- The **standalone .NET SDK** must also be installed (e.g.
  `winget install Microsoft.DotNet.SDK.8`). The SDK-style csproj
  needs `Microsoft.NET.Sdk` to resolve, and on Windows ARM64 the
  VS workload does not bundle it. After install, refresh PATH in
  any shell that was open before the install — see the deploy
  script for the canonical refresh idiom.
- Open `RBLclass.sln` in Visual Studio 2022. F5 launches Outlook with
  the add-in attached via *Debug → Start External Program:
  outlook.exe*. The HKCU COM registration written by the install
  script is what makes Outlook load the add-in; no VSTO hosting is
  involved.
- `RBLclass.Core` and its tests can be built and run from VS Code with
  the C# Dev Kit; the add-in shell project must be built from VS or
  the command line (`msbuild` or `dotnet build`) on Windows.
- Default build configuration is `Debug|AnyCPU`. Release is
  `Release|AnyCPU`. There is no x86 or x64 configuration.

## Publishing

The full release/packaging **workflow** (build → stage → verify both
native SQLite bitnesses → produce a shippable install kit) lives in the
manually-triggered **`/make-release`** skill at
`.claude/skills/make-release/` — run it to package the product for a
target workstation. The notes below are the surrounding facts; the runnable
steps are in the skill.

- Package format (current): a per-user PowerShell **install kit** (`.zip`)
  produced by `/make-release`, laying files under
  `%LocalAppData%\RBLclass\` and writing HKCU COM + Outlook Addins
  registry entries. No admin rights required.
- Package format (later, Phase 1 GA target): a WiX-based **MSI** under
  `/installer/`, per-user, optionally Authenticode-signed with the
  internal-PKI cert. Same registry/file effects as the install kit. The
  install kit and MSI coexist until the MSI is validated on the target.
- Signing is **optional** — Outlook does not require an Authenticode
  signature to load a COM add-in (validated in Phase 0).
- Distribution: internal HTTPS share documented in
  `docs/deployment.md`.

## Repository management

This is a solo project, developed interactively (with Claude) on a single
machine — which is also the only place changes can be built and tested.
The branching model is lightweight but strict about what may reach `main`.

### Branches

- `main` — always shippable. Updated **only** by merging `develop`. Never
  commit directly to `main`. Every commit on `main` is a feature or fix
  that has been validated locally and is ready to ship to the target.
- `develop` — integration branch and the default working branch. Direct
  commits are allowed (solo workflow). Always work on `develop` or a
  branch based on it.
- Feature branches — `feature/<slug>` or `fix/<slug>`, branched from
  `develop` and merged back into `develop`. Optional for small changes;
  expected for anything that might leave `develop` broken.

### Shipping a feature (`develop` → `main`)

Because this machine is the only test environment, nothing reaches `main`
until it has been built and validated here (and, for anything touching
COM / EDR / bitness, on the real 32-bit target — see the constraints
above).

1. Land the work on `develop`; confirm the solution builds and the
   `RBLclass.Core` tests pass.
2. Validate the behaviour interactively (run the add-in, exercise the
   feature).
3. Merge `develop` into `main` with a merge commit (`--no-ff`) so each
   ship is a distinct point in history.
4. Tag the merge on `main` with the product version `/make-release`
   stamps (e.g. `v1.0.0`).
5. Package from `main` with `/make-release` when an install kit is needed.

### GitHub (`gh` CLI)

- Use the `gh` CLI for all GitHub actions — PRs, releases, issues, repo
  settings. Do not drive GitHub through the web UI when a `gh` command
  exists.
- Pull requests are optional for solo merges; when used, target `develop`
  for features and `main` only for ship merges.
- Cut releases from tags on `main` (`gh release create vX.Y.Z`), attaching
  the install kit from `/make-release` when relevant.

### Hygiene

- Commit or push only when asked.
- Keep commits focused. Do not sweep unrelated working-tree changes into a
  commit (e.g. stray `.gitignore`, `docs/`, or `ROADMAP.md` edits) unless
  explicitly told to.
