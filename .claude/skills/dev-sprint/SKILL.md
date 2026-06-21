---
name: dev-sprint
description: Runs the full active sprint for the RBLclass add-in. Reads ROADMAP.md to find the highest vX.X.X section with unchecked items, then implements all items in document order using the feature-prep subagent and /feature-impl skill. Preparation for item N+1 starts in the background while item N is being implemented, so the brief is ready as soon as the next item begins.
---

# /dev-sprint — run the current sprint

## What this does

Finds the active sprint in `ROADMAP.md` and drives every unchecked item to
completion. Each item goes through:

1. **Preparation** (`feature-prep` subagent, Sonnet 4.6) — reads the
   codebase and produces an Implementation Brief.
2. **Implementation** (`/feature-impl` skill) — implements, verifies, and
   commits the item (including the ROADMAP.md checkbox update).

Preparation for item N+1 runs **in the background** while item N is being
implemented, so the brief is ready as soon as the next item begins.

---

## Rules

Always perform the dev sprint in develop branch, and make sure develop is up
to date with main before starting.

---

## Step 0 — Locate the active sprint

Read `ROADMAP.md`. The active sprint is the **highest-versioned** `## vX.X.X`
section that contains at least one unchecked item (`- [ ]`). Extract:
- The sprint version label (e.g. `v2.4.0.0`).
- The ordered list of unchecked items: label (A1, A2, B1, …), title, and
  full roadmap text.

If no unchecked items remain in any sprint section, report sprint complete
and stop.

---

## Step 1 — Prepare item[0]

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

## Step 2 — Pre-fetch item[1] in the background

Immediately after receiving the brief for item[0], if item[1] exists, start
its preparation without waiting for the result:

```
Agent({
  subagent_type: "feature-prep",
  run_in_background: true,
  description: "Prepare brief for [sprint] [next label]",
  prompt: "Sprint: [sprint version]\nItem: [next label] — [next title]\n\n[full next roadmap item text, verbatim from ROADMAP.md]"
})
```

---

## Step 3 — Implement item[0]

The brief for item[0] is in context. Invoke the implementation skill:

```
Skill("feature-impl")
```

`/feature-impl` handles build, tests, reload, user verification, commit, and
the ROADMAP.md checkbox. Wait for it to complete successfully (all
verification questions passed and commit made) before advancing.

---

## Step 4 — Advance to the next item

After item[0] is committed:
1. Wait for the background notification confirming item[1]'s brief is ready
   (if not already received).
2. If item[2] exists and has not yet been pre-fetched, start it in the
   background now (same pattern as Step 2).
3. The brief for item[1] is in context. Invoke `Skill("feature-impl")`.
4. Repeat until all items in the sprint are implemented and committed.

---

## Step 5 — Sprint complete

When the last item commits, report:
- Sprint version completed.
- All items done (label + title for each).
- Remind the user to run `/make-release` when ready to package, and to
  merge `develop` into `main` (with `--no-ff`) per CLAUDE.md once the
  build is validated on the 32-bit target.

---

## Rules

- **Never start item N+1 until item N is committed** (verification passed
  and commit made).
- **Never commit without passing verification** — enforced inside
  `/feature-impl`, honoured here too.
- If `/feature-impl` halts after 3 failed fix attempts, the sprint **pauses**.
  Surface the summary to the user and wait for guidance before continuing.
- The ROADMAP.md checkbox update is inside the feature commit (done by
  `/feature-impl`). The sprint skill never makes a standalone roadmap commit.
- If the user interrupts the sprint, resume by re-reading `ROADMAP.md` from
  scratch (Step 0) to discover which items remain unchecked.
