# Testing

RBLclass has a deliberately **split** validation strategy, because two of its
three layers can only be exercised inside a live Outlook process.

## What is automatically tested

`RBLclass.Core` — all the non-trivial business logic — is covered by **xUnit +
FluentAssertions + NSubstitute**. The highest-value surfaces are:

- the **folder-search matcher** (substring vs. word-prefix modes), and
- the **classify / send-guard decision logic**.

```powershell
dotnet test tests\RBLclass.Core.Tests
```

Because the core is `.NET Standard 2.0` with no Outlook, WPF, or COM
dependencies, the tests run on a plain runner with no Office installed — which is
exactly why CI can run them (see below).

### Resource parity is a test

`ResourceParityTests` (in `RBLclass.Core.Tests`) fails `dotnet test` if any
user-visible string key is **missing, empty, or has mismatched `{n}`
placeholders** across `Strings.resx` / `Strings.fr.resx` / `Strings.de.resx`, or
if a ribbon label/screentip/supertip is missing from any of the three
`Ribbon*.xml` variants. Localization is therefore enforced mechanically, not by
review.

## What is validated by hand

The **Outlook adapter** and the **add-in shell** need a live Outlook session and
real PST data, so they are excluded from automated tests and validated
interactively on the dev machine:

- Every user-visible increment ships an installable add-in. The `/reload-addin`
  loop rebuilds and reloads it into a running Outlook so the change can be
  checked live.
- Some behaviour **structurally cannot be exercised on the dev machine** — e.g.
  S/MIME (Stormshield) mail and the external-sender banner. Items in this
  category ship with an explicit *re-verify on the real target* note rather than
  being blocked.

## Target validation

Parity-critical and COM/EDR/bitness-sensitive increments are re-checked on the
**real 32-bit target** workstation before they reach pilots. The dev machine
(64-bit Outlook, ARM64 host under emulation, Current channel) is for fast
iteration only; final EDR and compatibility validation happens against a real
32-bit target.

## Continuous integration

`.github/workflows/ci.yml` runs the `RBLclass.Core` test suite on every push and
pull request to `develop` and `main`:

```yaml
- Setup .NET SDK 8.0.x
- dotnet test tests/RBLclass.Core.Tests/RBLclass.Core.Tests.csproj -c Release
```

Only Core (and its `net8.0` test project) is built in CI. The Outlook adapter and
add-in shell reference Office PIAs from the GAC, which a hosted runner does not
have, so they are intentionally excluded and validated interactively instead.

A separate workflow, `.github/workflows/docs.yml`, validates **this
documentation site** — see [Shipping → Documentation](shipping.md#documentation).

## The feature pipeline

Feature work runs through a three-part Claude-assisted pipeline that bakes
testing and verification in:

1. **`feature-prep`** (read-only subagent) turns a roadmap item into an
   Implementation Brief, including the exact **xUnit test cases** to add.
2. **`/feature-impl`** implements the brief, adds those tests, runs
   `dotnet test`, builds, reloads into Outlook, and then asks the human the
   brief's **verification questions** — a feature is never committed until every
   verification passes (or is explicitly marked not-verifiable-on-dev).
3. **`/dev-sprint`** orchestrates the two across a whole sprint.

The rule *"never commit a feature without a passing user verification"* is
enforced inside `/feature-impl`.
