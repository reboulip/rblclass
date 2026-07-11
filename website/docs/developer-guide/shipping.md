# Shipping

RBLclass is a solo project, developed interactively on a single machine — which
is also the only place changes can be built and tested. The branching model is
lightweight but strict about what may reach `main`.

## Branches

| Branch | Role |
| --- | --- |
| `main` | Always shippable. Updated **only** by merging `develop`. Never commit directly. Every commit is a validated, ready-to-ship change. |
| `develop` | Integration branch and default working branch. Direct commits allowed (solo workflow). |
| `feature/<slug>` / `fix/<slug>` | Branched from `develop`, merged back into `develop`. Optional for small changes; expected for anything that might leave `develop` broken. |

## Shipping a feature (`develop` → `main`)

Because this machine is the only test environment, nothing reaches `main` until
it has been built and validated here (and, for anything touching COM / EDR /
bitness, on the real 32-bit target):

1. Land the work on `develop`; confirm the solution builds and the
   `RBLclass.Core` tests pass.
2. Validate the behaviour interactively (run the add-in, exercise the feature).
3. Merge `develop` into `main` with a merge commit (`--no-ff`) so each ship is a
   distinct point in history.
4. Tag the merge on `main` with the product version (e.g. `v2.9.0.0`).
5. Package from `main` with `/make-release` when an install kit is needed.

!!! note "Version = COM strong-name"
    Changing `AssemblyVersion` changes the COM strong-name baked into the HKCU
    registration. Target workstations need a **reinstall** (re-run the install
    kit or MSI) to pick up a new version.

## Packaging

The full build → stage → verify → package workflow lives in the manually-invoked
`/make-release` skill. It produces two artifacts with identical registry/file
effects:

- **Install kit** (`.zip`, current) — a per-user PowerShell installer that lays
  files under `%LocalAppData%\RBLclass\` and writes HKCU COM + Outlook Addins
  entries. No admin rights.
- **MSI** (`/installer/`, WiX, GA target) — a per-user Windows Installer package
  pilots double-click. Same effects; coexists with the install kit until
  validated on the target.

A release **must** verify that both SQLite native bitnesses
(`runtimes\win-x86\native\` **and** `runtimes\win-x64\native\`) land in the
output — the single AnyCPU artifact has to load into both 32-bit and 64-bit
Outlook.

### Signing

Authenticode signing of the DLL and MSI with the internal-PKI Code Signing
certificate (EKU OID `1.3.6.1.5.5.7.3.3`), timestamped, is the GA target. Outlook
does **not** require a signature to load a COM add-in (validated in Phase 0), so
signing is currently optional. Expect a SmartScreen *"unrecognised publisher"*
warning on unsigned builds.

### Distribution

Internal HTTPS share (documented in `docs/deployment.md`). Releases are cut from
tags on `main` with `gh release create vX.Y.Z`, attaching the install kit from
`/make-release`.

## Documentation

Documentation ships **with the code, not after it**. This is a project rule, not
a nicety: every change that affects user-facing behaviour, architecture, the
tech stack, the testing/shipping process, or the security posture must include
the matching update to this site **in the same commit**.

- **Source:** `website/` (MkDocs Material). Pages are plain Markdown under
  `website/docs/`.
- **Who writes it:** the `docs-writer` subagent (Sonnet, medium reasoning
  effort) produces the update as part of the feature commit; `/feature-impl`
  dispatches it after verification passes. See
  [`CLAUDE.md` → Documentation](https://github.com/reboulip/rblclass/blob/main/CLAUDE.md).
- **CI:** `.github/workflows/docs.yml`
    - On a **pull request to `main`**, it builds the site with `--strict`, so a
      PR that breaks the docs build (bad link, broken nav) fails the check.
    - On **push to `main`** (i.e. once the PR merges), it builds and **publishes**
      the site to GitHub Pages.

!!! info "One-time Pages setup"
    The deploy job uses the GitHub Actions Pages flow
    (`actions/deploy-pages`). In the repository, set **Settings → Pages →
    Build and deployment → Source = GitHub Actions** once. No `gh-pages` branch
    is used.

### Building the docs locally

```powershell
python -m pip install -r website/requirements.txt
mkdocs serve -f website/mkdocs.yml     # live preview at http://127.0.0.1:8000
mkdocs build --strict -f website/mkdocs.yml   # what CI runs
```

## Hygiene

- Commit or push only when asked.
- Keep commits focused — don't sweep unrelated working-tree changes (`.gitignore`,
  stray `docs/`, `ROADMAP.md`) into a feature commit unless told to.
- Use the `gh` CLI for all GitHub actions (PRs, releases, issues).
