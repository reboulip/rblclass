---
name: dev-sprint
description: Runs the full active sprint for the RBLclass add-in. Reads ROADMAP.md to find the highest vX.X.X section with unchecked items, then implements all items in document order using the feature-prep subagent and /feature-impl skill. Preparation for item N+1 starts in parallel with N's implementation (same dispatch message); the brief is read and reconciled with actual changes once N's commit lands.
---

# /dev-sprint — run the current sprint

## What this does

Finds the active sprint in `ROADMAP.md` and drives every unchecked item to
completion. Each item goes through:

1. **Preparation** (`feature-prep` subagent, Sonnet 4.6) — reads the
   codebase and produces an Implementation Brief.
2. **Implementation** (`/feature-impl` skill) — implements, verifies, and
   commits the item (including the ROADMAP.md checkbox update).

Preparation for item N+1 is dispatched **in the same message** as `/feature-impl`
for item N, so both run in parallel. The brief lands while implementation is in
progress and is read right after N's commit, before N+1 starts.

---

## Rules

Always perform the dev sprint in develop branch, and make sure develop is up
to date with main before starting.

---

## Step 0 — Version stamp

### 0a — Locate the active sprint and read current version

Read `ROADMAP.md`. The active sprint is the **highest-versioned** `## vX.X.X.X`
section that contains at least one unchecked item (`- [ ]`). Extract the
sprint version label (e.g. `v2.4.0.0`).

Read the current `AssemblyVersion` from
`src/RBLclass.AddIn/Properties/AssemblyInfo.cs`.

### 0b — Classify the sprint and present a version proposal

Scan the sprint's full item list. Classify the overall change as:
- **Hotfix** — bug fixes only, no new user-facing capability.
- **Minor** — new features or enhancements; no breaking changes.
- **Major** — breaking changes, architectural overhaul, or a full-product
  version boundary.

Present a brief summary (5–10 words per item) and the classification
rationale, then ask the user to confirm or override the target version
via `AskUserQuestion`:

```
Sprint: <roadmap version>
Current AssemblyVersion: <x.y.z.w>
Classification: <Hotfix | Minor | Major> — <one-line rationale>

Planned changes:
  • <Label>: <brief description>
  …

Proposed version: <roadmap version>
```

Options: **"Confirm <roadmap version>"** / the user may select "Other" to
provide a different version string.

### 0c — Stamp version files (if needed)

If the confirmed version differs from the current `AssemblyVersion` in
`AssemblyInfo.cs`:

1. Update `src/RBLclass.AddIn/Properties/AssemblyInfo.cs`:
   - `[assembly: AssemblyVersion("X.X.X.X")]`
   - `[assembly: AssemblyFileVersion("X.X.X.X")]`
2. Update `.claude/skills/make-release/release.config.json`:
   - `"AssemblyVersion": "X.X.X.X"`
3. Commit both files:
   ```
   Bump version to <version>

   Co-Authored-By: Claude Sonnet 4.6 <noreply@anthropic.com>
   ```

If the versions already match, skip the file edits and the commit — just
state that the version is already correct and continue.

> **Note:** changing `AssemblyVersion` changes the COM strong-name baked
> into the HKCU registration. Target workstations will need a reinstall
> (re-run the install kit or MSI) to pick up the new version. Remind the
> user of this if the version changed.

---

## Step 1 — Locate the active sprint

Read `ROADMAP.md`. The active sprint is the **highest-versioned** `## vX.X.X.X`
section that contains at least one unchecked item (`- [ ]`). Extract:
- The sprint version label (e.g. `v2.4.0.0`).
- The ordered list of unchecked items: label (A1, A2, B1, …), title, and
  full roadmap text.

If no unchecked items remain in any sprint section, report sprint complete
and stop.

---

## Step 2 — Prepare item[0]

Spawn `feature-prep` for the first unchecked item (foreground — wait for
the result before continuing):

```
Agent({
  subagent_type: "feature-prep",
  description: "Prepare brief for [sprint] [label]",
  prompt: "Sprint: [sprint version]\nItem: [label] — [title]\n\n[full roadmap item text, verbatim from ROADMAP.md]"
})
```

---

## Step 3 — Implement item[0] and prepare item[1] in parallel

With item[0]'s brief in context, dispatch both in the **same message** so they
run concurrently:

1. Invoke the implementation skill (foreground — this call blocks until N is
   committed):
   ```
   Skill("feature-impl")
   ```
2. If item[1] exists, simultaneously spawn its preparation in the background:
   ```
   Agent({
     subagent_type: "feature-prep",
     run_in_background: true,
     description: "Prepare brief for [sprint] [next label]",
     prompt: "Sprint: [sprint version]\nItem: [next label] — [next title]\n\n[full next roadmap item text, verbatim from ROADMAP.md]"
   })
   ```

`/feature-impl` handles build, tests, reload, user verification, commit, and
the ROADMAP.md checkbox.

---

## Step 4 — Reconcile item[1]'s brief and advance

After item[0]'s commit lands:
1. Wait for the background notification confirming item[1]'s brief is ready
   (if not already received).
2. Review the brief against the codebase as it now stands after item[0]'s
   commit. Note any discrepancies — the brief was prepared against the
   pre-N codebase; item[0]'s changes may affect the context or affected
   files described in the brief. Carry those notes into `/feature-impl`.
3. Dispatch both in the **same message** (parallel):
   - Invoke `/feature-impl` for item[1] (foreground)
   - If item[2] exists and has not yet been pre-fetched, spawn item[2] prep
     in the background (same pattern as Step 3.2)
4. Repeat steps 4.1–4.3 for each subsequent item until all items are
   implemented and committed.

---

## Step 6 — Sprint complete

When the last item commits, report:
- Sprint version completed.
- All items done (label + title for each).
- Remind the user to run `/make-release` when ready to package (it will
  handle the develop → main PR, issue closing, tagging, and GH release).

---

## Rules

- **Never implement item N+1 until item N is committed** (verification passed
  and commit made). Background preparation of N+1 may run while N is being
  implemented, but `/feature-impl` for N+1 only starts after N's commit lands.
- **Never commit without passing verification** — enforced inside
  `/feature-impl`, honoured here too.
- If `/feature-impl` halts after 3 failed fix attempts, the sprint **pauses**.
  Surface the summary to the user and wait for guidance before continuing.
- The ROADMAP.md checkbox update is inside the feature commit (done by
  `/feature-impl`). The sprint skill never makes a standalone roadmap commit.
- If the user interrupts the sprint, resume by re-reading `ROADMAP.md` from
  scratch (Step 1) to discover which items remain unchecked. Skip Step 0
  (version stamp) on resume — the version was already committed at sprint start.
