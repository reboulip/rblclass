---
name: gh-issues
description: Triage open GH issues (implement / defer / drop with labels), then plan implement issues into ROADMAP.md and label/comment each one in GH.
---

# /gh-issues — plan open GH issues into the roadmap

## What this does

1. Fetches all open GH issues not yet carrying a disposition label (`roadmap`, `deferred`, or `dropped`).
2. Reads each one and classifies its nature.
3. **Triages each issue** with the user — implement, drop, or defer — and immediately applies the decided disposition (label + comment) before planning starts.
4. For issues marked *implement*: proposes a roadmap mapping, asks disambiguation questions as needed, writes items to `ROADMAP.md`, labels each issue `roadmap`, and posts a roadmap-link comment.

---

## Step 1 — Fetch open issues

```powershell
gh issue list --state open --json number,title,labels,body --limit 100
```

Filter out any issue that already carries one of the following labels — those are already triaged:
- `roadmap`
- `deferred`
- `dropped`

Also scan `ROADMAP.md` for `[#N]` back-references. If an issue number appears there but lacks the `roadmap` label, it is effectively planned — add the label silently and skip it from the triage flow.

**Re-review already-labelled open issues.** Labelled issues (`roadmap`,
`deferred`) that are still open are not necessarily settled: view each one
(`gh issue view <n> --json ...,comments,updatedAt`) and check for comments
added *after* the disposition was applied — new repro data, a fix that did
not hold, a follow-up suggestion. Any issue with substantive news re-enters
the flow: summarise what changed and ask the user (per issue, via
`AskUserQuestion`) whether to (re-)plan it, keep its current disposition, or
adjust labels. A `roadmap`-labelled issue whose planned fix shipped but
demonstrably failed gets a **new** roadmap item; a `deferred` issue with a
concrete new proposal may be promoted.

If no unresolved issues remain after filtering and re-review, report that and stop.

---

## Step 2 — Read issue details

For each unresolved issue:

```powershell
gh issue view <number> --json number,title,body,labels,comments
```

Classify each issue as one of:
- **Bug** — defect in existing functionality.
- **Enhancement** — improvement to existing functionality.
- **Feature** — new capability.

If a comment links a `github.com/user-attachments/files/...` attachment (e.g. a log
file from `%LocalAppData%\RBLclass\logs\`), fetch its raw content with
`curl -L -o <scratchpad-path> <url>` (or `gh api`), not `WebFetch` — WebFetch
summarizes plain-text attachments through a small model and can drop the exact
exception type/HRESULT/stack-trace lines needed to root-cause a bug.

Also form a tentative triage suggestion for each issue:
- Bugs → lean toward **implement** unless clearly out of scope.
- Enhancements/Features → weigh scope, complexity, and fit. Large or speculative items lean toward **defer**.

---

## Step 3 — Triage review

Present the full list of unresolved issues as a text summary with your suggested disposition:

```
#N — <title> [Bug / Enhancement / Feature]
   Suggested: implement / drop / defer
   Reason: <one sentence>

#M — <title> [Bug / Enhancement / Feature]
   Suggested: implement / drop / defer
   Reason: <one sentence>
```

Then use `AskUserQuestion` to collect the user's disposition for each issue. Process in batches of up to 4 issues per call (one question per issue), where each question has options:
- **Implement** — plan into roadmap
- **Defer** — label and set aside for a future cycle
- **Drop** — label and close; will not implement

After each batch, immediately apply the decided dispositions (Step 3a) before asking the next batch — this keeps progress visible and is resilient to the user stopping mid-review.

### Step 3a — Apply triage dispositions

For each issue just triaged, apply its disposition immediately:

**Dropped:**
```powershell
# Create label if needed
gh label create dropped --color b60205 --description "Will not be implemented" 2>$null

gh issue edit <number> --add-label dropped
gh issue comment <number> --body "Closing as out of scope — will not implement in the current product direction."
gh issue close <number>
```

**Deferred:**
```powershell
# Create label if needed
gh label create deferred --color e4e669 --description "Deferred to a future cycle" 2>$null

gh issue edit <number> --add-label deferred
gh issue comment <number> --body "Deferring for now — this may be revisited in a future cycle."
```

**Implement:** no action yet — these proceed to roadmap planning in Steps 4–8.

---

## Step 4 — Read ROADMAP.md

Read `ROADMAP.md` in full. Identify:
- The **active sprint section** (highest `## vX.X.X.X` with at least one unchecked item).
- Any open **future-phase sections** that could absorb new items.
- The **next available label** in each candidate section (to assign G1, H1, etc.).

Skip this step if no issues were triaged as *implement*.

---

## Step 5 — Propose a mapping

For each *implement* issue (or group of related issues), propose:
- **Target section**: an existing open section when it fits naturally; a new section when a different version makes more sense.
- **Item label**: next unused label in that section.
- **Item text**: one concise imperative sentence in the existing roadmap style.
- **Grouping rationale**: if two issues share a root cause or subsystem, propose merging them into one item; if one issue covers distinct sub-problems, propose splitting. Always explain the rationale.

When target section is uncertain, use `AskUserQuestion` before proceeding. Keep each question focused: present the proposed placement and ask the user to confirm or redirect. Do not batch all disambiguation into one giant question — ask per uncertain group.

**Ask implementation-approach questions.** For every *implement* issue with a
real design choice — where a fix should hook in, whether behaviour is
automatic or a setting, what a click/action does, how an edge case resolves —
put the choice to the user via `AskUserQuestion` (batches of up to 4, with a
recommended option first) *before* writing the roadmap. The answers are
decisions and must be baked into the item text; only skip an issue whose
implementation is genuinely unambiguous.

---

## Step 6 — Confirm the full plan

Before writing anything to ROADMAP.md, present the complete proposed mapping as a text summary:

```
Issue #N — <title>
→ <vX.X.X.X> <Label>: <proposed item text>

Issue #M — <title>
→ <vX.X.X.X> <Label>: <proposed item text>
```

Then use `AskUserQuestion`:
- **"Confirm — proceed"** — write the plan as proposed.
- **"Redirect — I'll provide corrections"** — wait for the user to describe what to change, then revise and ask again.

Only proceed to Step 7 after receiving explicit confirmation.

Skip this step if no issues are being implemented.

---

## Step 7 — Update ROADMAP.md

For each approved item, find the target section and append the item to its checklist:

```
- [ ] **<Label>.** <Item text>. [#N]
```

The `[#N]` at the end is the mandatory back-reference. If multiple issues map to one item: `[#N] [#M]`.

**Items are specifications, not summaries.** Each item must be a detailed,
self-contained spec that a different model (feature-prep/Sonnet), with no
access to this conversation, can implement efficiently: the observed
behaviour / repro facts from the issue thread, every implementation decision
the user made during triage, relevant prior-art references (earlier roadmap
items, commits, related settings), edge cases, and the localization/test
obligations the change triggers. Multi-sentence items are expected; a
one-liner is only acceptable for a genuinely trivial change.

If the target is a new version section, insert it after the active sprint section with an appropriate header, following the existing style (e.g. `## v2.5.0.0 — <theme>`).

After writing, verify that the ROADMAP.md structure is intact (no broken section hierarchy).

---

## Step 8 — Label implement issues in GH

```powershell
# Create the label if it does not exist
gh label create roadmap --color 0075ca --description "Planned in the roadmap" 2>$null

# Label each planned issue
gh issue edit <number> --add-label roadmap
```

---

## Step 9 — Post a planning comment on each implement issue

```powershell
gh issue comment <number> --body "This issue has been planned for development.

**Roadmap:** <vX.X.X.X> → <Label> — <item text>"
```

---

## Step 10 — Commit the planning result

Commit the run's changes on `develop` (authorised for this workflow — no
need to ask). Keep it to the files this skill touched: `ROADMAP.md`, plus
`docs/roadmap-history.md` and/or skill files if the run updated them. Use a
message like `Plan GH issues #N, #M into ROADMAP (vX.X.X.X)`; put skill/doc
process changes in a separate focused commit. Do not push unless asked.

---

## Rules

- Never write to `ROADMAP.md` before Step 6 confirmation.
- Always include a `[#N]` back-reference on every roadmap item sourced from a GH issue.
- Apply triage dispositions (Step 3a) per batch as soon as the user confirms — do not wait until the end of the full triage.
- `dropped` issues are closed in GH; `deferred` issues remain open.
- If a proposed new version section could conflict with existing version numbering, ask the user before creating it.
- Unrelated issues can live in the same release section. A hotfix section before a minor section is valid when the urgency warrants it — explain the reasoning when proposing it.
