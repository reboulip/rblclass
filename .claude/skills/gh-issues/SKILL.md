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

If no unresolved issues remain after filtering, report that and stop.

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

## Rules

- Never write to `ROADMAP.md` before Step 6 confirmation.
- Always include a `[#N]` back-reference on every roadmap item sourced from a GH issue.
- Apply triage dispositions (Step 3a) per batch as soon as the user confirms — do not wait until the end of the full triage.
- `dropped` issues are closed in GH; `deferred` issues remain open.
- If a proposed new version section could conflict with existing version numbering, ask the user before creating it.
- Unrelated issues can live in the same release section. A hotfix section before a minor section is valid when the urgency warrants it — explain the reasoning when proposing it.
