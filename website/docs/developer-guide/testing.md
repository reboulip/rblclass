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

The experimental re-encrypted filing is covered at the Core level by
`MimeEntityBuilderTests` (with the `MimeTestReader` helper),
`ReencryptContentPlannerTests`, `EncryptionCertificatePolicyTests`,
`ReencryptEligibilityTests`, `ClassifierServiceReencryptTests` and
`CmsEnvelopeBuilderTests` (parses the envelope with `System.Formats.Asn1` from the
net8.0 shared framework, decrypts it with the private key, checks a fresh key/IV
per envelope and minimal INTEGER serials). The Save-as
planner (`SaveAsPlannerTests`) and the phishing-warning decisions
(`PhishingLinkTests`, `PhishingWarningPolicyTests`) are Core logic with their own
tests; the hooks, Word automation and dialogs around them are not.

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
- **Re-encrypted filing** is exercised end to end on the dev machine with a
  self-signed `CN=RBLclass E2E Test` certificate (no SAN e-mail, so *Automatic*
  does not select it and the tests pick it explicitly) and a **simulated
  Stormshield decryption**: the item is stored as `IPM.Note.SMIME` while an
  external COM handle holds it with an unsaved in-memory class `IPM.Note` during
  the classify. The log shows key transport `1.2.840.113549.1.1.1`, and the copy
  is stored `IPM.Note.SMIME` and opens decrypted with all its parts. Real
  Stormshield behaviour is verified only on the target.
- **Store-level filing of protected mail** (`MapiMessageMover`) is covered in Core
  by `ClassifierServiceStoreLevelTests` (untracked and verified refs). The MAPI
  call itself is exercised end to end on the dev machine with the same simulated
  Stormshield decryption and a real Stormshield `smime.p7m` envelope; the real
  card-inserted case on 32-bit Outlook is verified only on the target.
- The **phishing warnings** and **PDF export** are exercised end to end on the dev
  machine by driving Outlook through UI Automation and COM with fake mails that
  carry a simulated banner (link click, Enter, context menu, attachment open,
  Cancel and *Open anyway*). What stays real-target only: a genuine banner or
  Stormshield mail, the corporate EDR's reaction to the thread-scoped
  `SetWindowsHookEx` calls on 32-bit Outlook, and Word automation under the EDR.
- `tools/Inspect-SmimeMsg.ps1` reads a dragged-out `.msg` offline and reports its
  S/MIME CMS envelope (content-encryption and key-transport algorithms, and the
  recipients).

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

Feature work runs through the generic **`/dev-pipeline`** skill (global,
`~/.claude/skills/`) and issue triage through **`/gh-issues`** (also global).
The only project-local skills are `/make-release` and `/reload-addin`.

A sprint runs in dependency-ordered **waves**:

1. A `sprint-planner` subagent groups roadmap items into waves at the start
   of the sprint.
2. `feature-forecast` subagents produce Implementation Briefs (including the
   **xUnit test cases** to add) one wave ahead, in parallel with the previous
   wave's implementation.
3. Each wave is implemented, then tested with `/test-select` (`dotnet test`).
4. **Per-wave gate** — before the wave is committed:
    - Build the add-in with MSBuild.
    - Run `/reload-addin` to reload it into a live Outlook session.
    - Ask the wave's **verification questions** via `AskUserQuestion`; a wave
      is never committed until every verification passes.
5. The wave lands as **one commit** on a `feature/<milestone-slug>` branch off
   `develop`, merged back with `--ff-only`.

**Not verifiable on the dev machine.** Items that depend on receiving
S/MIME (Stormshield) mail or an external-sender banner are committed with the
checkbox left open and a `- shipped, pending tests` note, and the linked issue
gets the `awaiting-verification` label. Everything else in the same item still
goes through the normal verification questions.
