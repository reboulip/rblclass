# Tech stack

RBLclass is a **classic Outlook Shared COM Add-in** — the pre-VSTO,
`IDTExtensibility2` + `IRibbonExtensibility` model — deliberately chosen so it
can be installed per-user without admin rights, work fully offline against PST
files, and load on the classic Win32 Outlook client.

!!! info "Why not VSTO?"
    VSTO is explicitly excluded. The Office Developer Tools workload could not be
    set up on the dev machine, and the Shared COM Add-in model covers everything
    the project needs (PST access, offline use, per-user install). ClickOnce is
    not used either — it is VSTO-specific in this scenario.

## Runtimes and language

| Component | Target framework | Why |
| --- | --- | --- |
| `RBLclass.AddIn` (shell) | .NET Framework 4.8 | Hosts COM interop, WPF, ribbon, task panes. |
| `RBLclass.Outlook.Adapter` | .NET Framework 4.8 | COM access to the Outlook Object Model. |
| `RBLclass.Core` (business logic) | .NET Standard 2.0 | Portable, Outlook-free core. |

- **Language:** C# 7.3 (the .NET Framework 4.8 default; `LangVersion` is
  deliberately not bumped).
- **Build configuration:** **AnyCPU** for all projects, with
  `Prefer32Bit=false`. The same binary JITs into 64-bit Outlook on the ARM64 dev
  machine (running as x64 under Prism emulation) **and** 32-bit Outlook on the
  x86 target workstations. There is no x86 or x64 configuration.

## Key components

| Concern | Technology |
| --- | --- |
| UI | WPF with a **hand-rolled minimal MVVM** (`ObservableObject` + `RelayCommand`), hosted in Custom Task Panes via `ElementHost` inside a ComVisible WinForms host. |
| Ribbon | `IRibbonExtensibility.GetCustomUI` returning Ribbon XML (not the Ribbon Designer). |
| Storage | SQLite via **Microsoft.Data.Sqlite**, with **FTS5** for full-text folder search. Database in `%LocalAppData%\RBLclass\`. |
| Logging | **Serilog** with a rolling file sink under `%LocalAppData%\RBLclass\logs\`. |
| Tests | **xUnit + FluentAssertions + NSubstitute** against `RBLclass.Core`. |
| Packaging | Per-user PowerShell install kit (`.zip`) now; per-user **WiX MSI** at GA, both writing HKCU + `%LocalAppData%`. |

## Deliberate dependency constraints

The COM host has **no binding redirects**, so every transitive strong-named
dependency version must match exactly what the SQLite facade assemblies were
built against — otherwise the first `SqliteConnection` throws at runtime.

!!! warning "CommunityToolkit.Mvvm is deliberately NOT used"
    Its 8.x dependencies (`System.Memory` ≥ 4.5.5,
    `System.Runtime.CompilerServices.Unsafe` ≥ 6.0.0) float the SQLite facade
    assemblies above what `SQLitePCLRaw` was built against, and with no binding
    redirects that breaks the first `SqliteConnection`. This is why the MVVM
    layer is hand-rolled and why new AddIn dependencies must stay
    dependency-light.

Two more standing rules that shape the stack:

- **Native SQLite must ship both bitnesses.** `SQLitePCLRaw.bundle_e_sqlite3`
  payloads must land under both `runtimes\win-x86\native\` and
  `runtimes\win-x64\native\`. **No x64-only dependencies.**
- **Memory budget ~1 GB.** On the 32-bit deployment target the process shares a
  ~2 GB address space with Outlook itself. Stream from SQLite; don't load large
  datasets into memory.

## Development prerequisites

- **Visual Studio 2022** with the **.NET desktop development** workload and the
  **.NET Framework 4.8 targeting pack**. The Office/SharePoint workload is *not*
  required (but Office must be installed locally for the PIA references).
- The **standalone .NET SDK** (e.g. `winget install Microsoft.DotNet.SDK.8`) —
  the SDK-style `.csproj` needs `Microsoft.NET.Sdk`, which the VS workload does
  not bundle on Windows ARM64.
- `RBLclass.Core` and its tests build and run from VS Code with the C# Dev Kit;
  the add-in shell must be built from VS or the command line (`msbuild` /
  `dotnet build`) on Windows.

## Localization

The UI is English / French / German, following Outlook's UI language (defaulting
to English) with a per-user override resolved once at startup. Every user-visible
string is a resource key in `Strings.resx` + `Strings.fr.resx` + `Strings.de.resx`;
`ResourceParityTests` fails the build if a key is missing, empty, or has
mismatched `{n}` placeholders in any language. See
[Architecture](architecture.md#localization) for how strings flow through XAML and
code.
