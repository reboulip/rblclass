---
name: feature-impl
description: Implements one RBLclass roadmap item end-to-end. Expects an Implementation Brief in the current context (produced by the feature-prep subagent, typically via /dev-sprint). If no brief is in context, invokes feature-prep first. After implementing, runs /reload-addin and asks the user the brief's verification questions via AskUserQuestion. On all-pass: updates the ROADMAP.md checkbox and commits. On any failure: diagnoses and fixes before re-verifying.
---

# /feature-impl — implement one sprint item end-to-end

## Prerequisites

An **Implementation Brief** (produced by the `feature-prep` subagent) must be
in the current conversation context. If one is not present, invoke it now:

```
Agent({
  subagent_type: "feature-prep",
  description: "Prepare implementation brief",
  prompt: "Sprint: [version]\nItem: [label] — [title]\n\n[full roadmap item text]"
})
```

Then continue once the brief is returned.

## Step 1 — State the item

Before touching any file, state clearly which item you are implementing
(e.g. "Implementing v2.4.0.0 A1 — Clear the search field after classify").
This anchors the conversation.

## Step 2 — Implement

Work through each file in the brief's "Files to modify" and "Files to create"
sections. Read the current file content (or the relevant section) before
editing. Apply exactly the changes described in "Required changes" — no
opportunistic refactoring of surrounding code.

Follow CLAUDE.md coding rules throughout: no comments unless the WHY is
non-obvious, no extra abstractions, no error handling for impossible cases.

## Step 3 — Localization

For every key listed in the brief's "Localization" section:
- Add the key + English value to `src/RBLclass.AddIn/Resources/Strings.resx`.
- Add translated values to `Strings.fr.resx` and `Strings.de.resx`.

`ResourceParityTests` fails the build if any key is missing or has mismatched
`{n}` placeholders — never skip this step.

## Step 4 — xUnit tests

For every test case in the brief's "xUnit tests" section:
- Add it to the specified test class in `tests/RBLclass.Core.Tests/`.
- Run `dotnet test tests/RBLclass.Core.Tests/ --verbosity minimal` and confirm
  all tests pass before proceeding.

## Step 5 — Build check

`msbuild` is not on PATH on this machine — resolve it via `vswhere` first
(same lookup `Reload-Addin.ps1` uses), then build to catch compilation
errors before reloading into Outlook:

```powershell
$vswhere = "${env:ProgramFiles(x86)}\Microsoft Visual Studio\Installer\vswhere.exe"
$msbuild = & $vswhere -latest -prerelease -products * -requires Microsoft.Component.MSBuild -find "MSBuild\**\Bin\MSBuild.exe" | Select-Object -First 1
if (-not $msbuild -or -not (Test-Path $msbuild)) { throw "MSBuild not found via vswhere." }
& $msbuild "src\RBLclass.AddIn\RBLclass.AddIn.csproj" /p:Configuration=Debug "/p:Platform=Any CPU" /t:Build /v:minimal /nologo
```

Fix any errors before continuing. Do not call /reload-addin until the build
is clean.

## Step 6 — Reload

```
Skill("reload-addin")
```

This closes Outlook, installs the fresh build, and restarts Outlook. Tell the
user what changed and what to look for when Outlook reopens.

## Step 7 — Verify

Present all verification questions from the brief's "Verification questions"
section to the user via AskUserQuestion (all at once, max 4). For each
question offer at minimum: `Yes — works as expected` and
`No — something is wrong`.

Always end with a last question: `Any other feedback?` with a free text response, to catch anything the user notices that the specific verification questions missed.

Wait for the answers before doing anything else.

**Not verifiable on the dev machine.** CLAUDE.md documents a category of
behavior that structurally cannot be exercised on this machine (e.g.
S-MIME/Stormshield mail, the external-sender banner). If the user reports a
question can't be answered for this reason, don't treat it as a fail and
don't block the sprint on it — follow Step 8c instead.

## Step 8a — All pass → commit

1. Edit `ROADMAP.md`: find the item's checkbox line in the active sprint
   section (matches `- [ ] **[Label].`, e.g. `- [ ] **A1.`) and change
   `[ ]` to `[x]`.
2. Stage all changed files: source, tests, localization resources, ROADMAP.md.
3. Commit with a message following the existing commit style:
   ```
   [Imperative summary] ([sprint version] [label])

   Co-Authored-By: Claude Sonnet 4.6 <noreply@anthropic.com>
   ```
   Example: `Clear search field after classify (v2.4.0.0 A1)`
4. Do **not** push — that is a separate user action.
5. **GH issue comment (if applicable).** Scan the roadmap item's full text
   for `[#N]` back-references (regex `\[#(\d+)\]`). If any are found:
   ```powershell
   $sha = git rev-parse HEAD
   $shortSha = git rev-parse --short HEAD
   $repoUrl = gh repo view --json url --jq .url
   gh issue comment <N> --body "Implemented in commit [$shortSha]($repoUrl/commit/$sha).
   Forecasted release: <sprint version>."
   ```
   The sprint version is the `## vX.X.X.X` section header of the item being
   implemented. If no `[#N]` ref is present, skip this step silently.

## Step 8b — Any fail → diagnose and loop

1. Ask one focused follow-up question if the symptom needs clarification.
2. Diagnose the root cause from the user's description and the code.
3. Apply the fix.
4. Return to Step 5 (build check) and repeat from there.

After **3 failed fix attempts**, stop the loop. Summarise what was tried,
what each attempt changed, and what the user reported. Ask the user for
guidance before attempting another fix.

## Step 8c — Not verifiable on the dev machine → commit unverified

Follow the precedent already in `ROADMAP.md` (v2.6.0.0 B3, C1; v2.9.0.0 B1):

1. Same as Step 8a, except leave the checkbox as `- [ ]` instead of `- [x]`.
2. Append `- shipped, pending tests (not verifiable on the dev machine per
   the testing strategy above; re-verify <what to check> on the real target
   before pilot)` to the item's text, naming the specific behavior that
   still needs a real-target check.
3. Commit and comment on the GH issue as in Step 8a, but note in the issue
   comment that verification is pending real-target testing.

## Rules

- Never commit until **all** verification questions pass.
- Never commit without including the ROADMAP.md `[ ]` → `[x]` update.
- Never push to remote — push is a separate user action.
- xUnit must pass (Step 4) before calling /reload-addin (Step 6).
- Build must succeed (Step 5) before calling /reload-addin (Step 6).
