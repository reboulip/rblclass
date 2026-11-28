# CLAUDE.md

## Interaction language

The project is in English: code, identifiers, comments, commit messages,
ROADMAP items, documentation and every English resource string. **Talk to the
user in French** (replies, questions, `AskUserQuestion` prompts, summaries).
The French and German UI strings stay in their resx files as they are.

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

### Localization (English / French / German)

The UI follows Outlook's UI language, defaulting to English for any other
language, with a per-user override in Settings resolved once at startup
(`RblClassAddIn.OnStartupComplete`, before any pane/window/ribbon is
created). Changing the "Preferred UI language" setting takes effect after
restarting Outlook.

- Every new user-visible string MUST be added as a resource key to
  `src/RBLclass.AddIn/Resources/Strings.resx` **and** translated in
  `Strings.fr.resx` and `Strings.de.resx` in the same change.
  `ResourceParityTests` (in `RBLclass.Core.Tests`) fails `dotnet test` if a
  key is missing, empty, or has mismatched `{n}` placeholders in any
  language.
- XAML strings go through `{loc:Loc Key=...}`
  (`RBLclass.AddIn.Localization.LocExtension`). Code-behind/ViewModel
  strings go through `ILocalizationService.GetString`/`GetString(key, args)`
  (`TaskPaneServices.Localization`).
- Count-dependent strings always get a `_One` / `_Other` resource key pair,
  resolved via `ILocalizationService.Plural(count, oneKey, otherKey, args)`.
  Never derive plurals by string concatenation — French and German plural
  forms differ from English.
- New ribbon buttons/tooltips need matching `label`/`screentip`/`supertip`
  entries in `Ribbon.xml`, `Ribbon.fr.xml`, and `Ribbon.de.xml` — checked by
  `ResourceParityTests` as well. `RblClassAddIn.GetCustomUI` picks the
  variant based on `TaskPaneServices.Localization.CurrentLanguage`.

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

## Documentation

The product has an online documentation site published to GitHub Pages,
built with **MkDocs Material**. Its source is `website/` (Markdown pages
under `website/docs/`, navigation in `website/mkdocs.yml`). It has four
areas — **Home**, **User Guide** (installation, quick start, walkthrough),
**Developer Guide** (tech stack, architecture, testing, shipping), and
**Security**.

**Documentation ships with the code, not after it.** Every change that
alters any of the following MUST be accompanied by the matching docs update
— **in the same commit** for a change made outside a sprint, and in the
sprint's single docs commit for a change made by `/dev-pipeline` (see below):

- user-visible behaviour, a setting, a default, or a keyboard shortcut →
  `website/docs/user-guide/` (the walkthrough settings tables are the
  source of truth) and, if a feature is added/removed/renamed, the Home
  feature bullets;
- the architecture, a coding rule, the tech stack, a dependency, or a
  build/test/ship process → `website/docs/developer-guide/`;
- the security posture (new data persisted, a network call, a new trust
  boundary, a closed hardening item) → **both** `website/docs/security.md`
  (published) and `docs/security.md` (canonical internal threat model).

Do not hand-write these updates inline: delegate them to the global
**`doc-keeper` subagent**. During a sprint, `/dev-pipeline` hands it each
committed wave and lands its work in one docs-only commit before the merge.
For a manual change outside a sprint, give it the change description and the
diff, then stage its output into the same commit as the change.

### Documentation Ownership

The docs directories are `website/docs/` and `docs/security.md`, with nav in
`website/mkdocs.yml`. A new page must be added to `nav:` in the same edit.

| Page | Covers | Update it when… |
|------|--------|-----------------|
| `website/docs/index.md` | Landing page: pitch, feature bullets, requirements at a glance | A feature is added/removed/renamed, or a top-level requirement changes |
| `website/docs/user-guide/installation.md` | Prerequisites, install, first run, data locations, update, uninstall | Install/uninstall flow, prerequisites, data paths, or the update mechanism change |
| `website/docs/user-guide/quick-start.md` | The two-minute "file your first email" path | The core classify/search flow or its default keys change |
| `website/docs/user-guide/walkthrough.md` | Full feature and **settings reference** tables, keyboard shortcuts | Any user-visible behaviour, setting, default, or shortcut changes |
| `website/docs/developer-guide/tech-stack.md` | Runtimes, key components, dependency constraints, dev prerequisites | The stack, a dependency, a framework target, or a build prerequisite changes |
| `website/docs/developer-guide/architecture.md` | Layering, COM/lifetime/threading rules, storage/SQLite, events, perf targets | The architecture, a coding rule, the schema policy, or a perf target changes |
| `website/docs/developer-guide/testing.md` | What's auto-tested vs. hand-validated, CI, the feature pipeline | The test strategy, CI, or resource-parity enforcement changes |
| `website/docs/developer-guide/shipping.md` | Branching, ship flow, packaging, signing, distribution, the docs pipeline | The branching/release/packaging/signing process changes |
| `website/docs/security.md` | Trust boundaries, data handled, network posture, accepted-by-design, hardening backlog | The **security posture** changes (new data persisted, a network call, a new trust boundary, a hardening item closed) |

- The **walkthrough settings tables** are the source of truth for user-facing
  settings and defaults. When one changes, update the exact table row and check
  whether `index.md` or `quick-start.md` also mention it.
- A security-posture change goes in BOTH `website/docs/security.md` and the
  canonical internal `docs/security.md`, kept consistent.
- The developer pages paraphrase this file: mirror a changed rule in the matching
  page, and do not invent rules that are not here.
- Match the existing voice: concise, second person for user docs, present tense,
  Material admonitions (`!!! note`, `!!! warning`, `!!! danger`) where the pages
  already use them. Prefer extending a page over adding one.

CI (`.github/workflows/docs.yml`) builds the site with `--strict` on every
pull request to `main` (a broken link or a page missing from `nav:` fails
the check) and **publishes** to GitHub Pages on push to `main` (i.e. once
the PR merges). Keep internal links valid so the strict build passes.

## Sprint development workflow

Sprints run through the generic **`/dev-pipeline`** skill and issue triage
through **`/gh-issues`**, both from the global config
(`~/.claude/skills/`, repo `reboulip/claude-config`). The only project skills
are `/make-release` and `/reload-addin`. The generic skills defer to
`CLAUDE.md` for project conventions; these are RBLclass's, and they override
the generic defaults wherever the two disagree.

### ROADMAP conventions

- Sprint milestones are `## vX.X.X.X — <theme>` sections, with items
  `- [ ] **A1.** …` grouped under `### A — <topic>` headings.
- **The active milestone is the highest-versioned `## vX.X.X.X` section with
  an unchecked item**, not the first open one. Older sections keep unchecked
  items tagged `- shipped, pending tests`: those await a real-target re-check
  and are out of sprint scope, as are `## Phase N` sections and the open
  carried-over items.
- An item is a **self-contained detailed spec**, implementable by a smaller
  model without the planning conversation — not a one-liner. Dependencies use
  `(requires: <label>)` and GitHub issues `[#N]`.

### Toolchain

- **Test command:** `dotnet test tests/RBLclass.Core.Tests/ --verbosity minimal`.
  Full suite: under 1s of test time for 344 tests, measured 2026-11-05, so
  always run all of it.
- **Formatter:** none is enforced (`.editorconfig` only, no format CI gate);
  skip the pipeline's format steps.
- **Sprint branch:** `feature/<milestone-slug>` off `develop`, merged back with
  `--ff-only`. The integration branch is `develop`, the release branch `main`.
- **Release:** `/make-release`, not the generic `/release`.
- **Docs build:** `mkdocs build --strict -f website/mkdocs.yml` (CI runs it; run
  it locally only if `mkdocs` is installed).

### Sprint start: version stamp

Before branching, confirm the target version with the user, then update
`AssemblyVersion` and `AssemblyFileVersion` in
`src/RBLclass.AddIn/Properties/AssemblyInfo.cs`, `Version`, `AssemblyVersion`
and `FileVersion` in `src/RBLclass.Core/RBLclass.Core.csproj` and
`src/RBLclass.Outlook.Adapter/RBLclass.Outlook.Adapter.csproj`, and
`AssemblyVersion` in `.claude/skills/make-release/release.config.json`, and
commit `Bump version to <version>`. Every RBLclass DLL carries the product
version because the MSI replaces an installed DLL only when its file version
changes: a DLL left on an old version keeps its previous build on upgrade.
`VersionParityTests` fails `dotnet test` and `/make-release` refuses to
package on a mismatch. A changed `AssemblyVersion` changes the COM strong name,
so target workstations need a reinstall: remind the user. Skip this on a
resumed sprint.

### Per-wave gate, between the green tests and the wave commit

**Never commit a wave without a passing user verification.** After
`/test-select` is green:

1. Build the add-in. `msbuild` is not on PATH, and the lookup must require the
   SDK component or it picks a Build Tools instance that cannot resolve
   `Microsoft.NET.Sdk`:
   ```powershell
   $vswhere = "${env:ProgramFiles(x86)}\Microsoft Visual Studio\Installer\vswhere.exe"
   $msbuild = & $vswhere -latest -prerelease -products * -requires Microsoft.Component.MSBuild Microsoft.NetCore.Component.SDK -find "MSBuild\**\Bin\MSBuild.exe" | Select-Object -First 1
   & $msbuild "src\RBLclass.AddIn\RBLclass.AddIn.csproj" /p:Configuration=Debug "/p:Platform=Any CPU" /t:Build /v:minimal /nologo
   ```
2. Run `/reload-addin`, warning the user that it closes Outlook, and say what
   to look for.
3. Ask the wave's verification questions through `AskUserQuestion` (at most 4
   per call, each offering at least `Yes — works as expected` and
   `No — something is wrong`), ending with a free-text `Any other feedback?`.
4. On any failure, diagnose, fix and repeat from step 1. After 3 failed fix
   attempts, pause and ask the user for guidance.

**Not verifiable on the dev machine.** Never ask the user to live-test anything
that depends on receiving S/MIME/Stormshield mail or an external-sender banner:
this machine never gets such mail. Commit those parts anyway, but leave the
checkbox `- [ ]` and append `- shipped, pending tests (not verifiable on the dev
machine per the testing strategy above; re-verify <what to check> on the real
target before pilot)`, then add the `awaiting-verification` label to the issue.
Anything else in the same item still goes through the normal questions.

### After each wave commit

For every `[#N]` in a committed item, post `Implemented in commit <sha link>.
Forecasted release: <version>.` with `gh issue comment`, noting any pending
real-target verification. Do not push: that remains a separate user action.

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

- When committing from PowerShell, always write the message to a temp file
  and use `git commit -F <file>` — PowerShell 5.1 mangles embedded double
  quotes in `-m` here-strings passed to native executables.
- Commit or push only when asked.
- Keep commits focused. Do not sweep unrelated working-tree changes into a
  commit (e.g. stray `.gitignore`, `docs/`, or `ROADMAP.md` edits) unless
  explicitly told to.
