---
name: gh-issues
description: Review open GH issues that are not yet planned, propose roadmap items for each, ask disambiguation questions as needed, then write them to ROADMAP.md and label/comment each issue in GH.
---

# /gh-issues — plan open GH issues into the roadmap

## What this does

1. Fetches all open GH issues not yet carrying the `roadmap` label.
2. Reads each one's full details and classifies its nature.
3. Proposes a mapping to roadmap items — which version section, how to group or split, and the item text — asking disambiguation questions where placement is uncertain.
4. After explicit user confirmation: writes items to `ROADMAP.md`, labels each issue `roadmap` in GH, and posts a roadmap-link comment on each issue.

---

## Step 1 — Fetch open issues

```powershell
gh issue list --state open --json number,title,labels,body --limit 100
```

Filter out any issue that already carries the `roadmap` label — those are planned.

Also scan `ROADMAP.md` for `[#N]` back-references. If an issue number appears there but lacks the label, it is effectively planned — add the label silently and skip it from the planning flow.

---

## Step 2 — Read issue details

For each unplanned issue:

```powershell
gh issue view <number> --json number,title,body,labels,comments
```

Classify each issue as one of:
- **Bug** — defect in existing functionality. Candidate for the active sprint or a hotfix section.
- **Enhancement** — improvement to existing functionality. Candidate for the next minor version.
- **Feature** — new capability. Candidate for the next minor or major version.

---

## Step 3 — Read ROADMAP.md

Read `ROADMAP.md` in full. Identify:
- The **active sprint section** (highest `## vX.X.X.X` with at least one unchecked item).
- Any open **future-phase sections** that could absorb new items.
- The **next available label** in each candidate section (to assign G1, H1, etc.).

---

## Step 4 — Propose a mapping

For each unplanned issue (or group of related issues), propose:
- **Target section**: an existing open section when it fits naturally; a new section when a different version makes more sense (e.g. a hotfix before a planned minor, or a clearly forward-looking feature).
- **Item label**: next unused label in that section.
- **Item text**: one concise imperative sentence in the existing roadmap style.
- **Grouping rationale**: if two issues share a root cause or subsystem, propose merging them into one item; if one issue covers distinct sub-problems, propose splitting. Always explain the rationale.

When target section is uncertain, use `AskUserQuestion` before proceeding. Keep each question focused: present the proposed placement and ask the user to confirm or redirect. Do not batch all disambiguation into one giant question — ask per uncertain group.

---

## Step 5 — Confirm the full plan

Before writing anything, present the complete proposed mapping as a text summary:

```
Issue #N — <title>
→ <vX.X.X.X> <Label>: <proposed item text>

Issue #M — <title>
→ <vX.X.X.X> <Label>: <proposed item text>
```

Then use `AskUserQuestion`:
- **"Confirm — proceed"** — write the plan as proposed.
- **"Redirect — I'll provide corrections"** — wait for the user to describe what to change, then revise and ask again.

Only proceed to Step 6 after receiving explicit confirmation.

---

## Step 6 — Update ROADMAP.md

For each approved item, find the target section and append the item to its checklist:

```
- [ ] **<Label>.** <Item text>. [#N]
```

The `[#N]` at the end is the mandatory back-reference. If multiple issues map to one item: `[#N] [#M]`.

If the target is a new version section, insert it after the active sprint section with an appropriate header, following the existing style (e.g. `## v2.5.0.0 — <theme>`).

After writing, verify that the ROADMAP.md structure is intact (no broken section hierarchy).

---

## Step 7 — Label issues in GH

```powershell
# Create the label if it does not exist
gh label create roadmap --color 0075ca --description "Planned in the roadmap" 2>$null

# Label each planned issue
gh issue edit <number> --add-label roadmap
```

---

## Step 8 — Post a planning comment on each issue

```powershell
gh issue comment <number> --body "This issue has been planned for development.

**Roadmap:** <vX.X.X.X> → <Label> — <item text>"
```

---

## Rules

- Never write to `ROADMAP.md` before Step 5 confirmation.
- Always include a `[#N]` back-reference on every roadmap item sourced from a GH issue.
- Never close or resolve GH issues — this skill only plans them.
- If a proposed new version section could conflict with existing version numbering, ask the user before creating it.
- Unrelated issues can live in the same release section. A hotfix section before a minor section is valid when the urgency warrants it — explain the reasoning when proposing it.
