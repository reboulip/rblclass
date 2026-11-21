# RBLclass roadmap

**Closed history** — Phases 0–7 and every shipped version (**v2.0.0.0**
through **v2.12.0.0**) — lives in
[docs/roadmap-history.md](docs/roadmap-history.md). This file holds **where we
are**, the **open carried-over items**, the next sprint milestone once one is
planned, and the **forward-looking phases**.

Reference docs from the legacy analysis:
[docs/legacy-overview.md](docs/legacy-overview.md),
[docs/legacy-reimplementation-roadmap.md](docs/legacy-reimplementation-roadmap.md),
[docs/parity-and-regression.md](docs/parity-and-regression.md).

## Where we are (2026-11-21)

The rewritten add-in reached full legacy parity as **v2.0.0.0**, went through
the pilot and general availability (Phases 6–7), and has since shipped
feedback waves up to **v2.12.0.0** (experimental re-encrypted filing of
protected mail, 2026-11-21 —
[release](https://github.com/reboulip/rblclass/releases/tag/v2.12.0.0)). All
of them are recorded in [history](docs/roadmap-history.md). No sprint
milestone is planned yet: the next one comes from a `/gh-issues` triage.

Open work: the two real-target items below, encrypted (Stormshield) mail
[#62] and the external-sender banner on reply/forward [#46]. Phases 8–9 are
forward-looking.

## Testing & validation strategy (applies to every item)

- **Automated unit tests:** all non-trivial logic lives in `RBLclass.Core`,
  covered by **xUnit + FluentAssertions + NSubstitute**. The folder-search
  matcher and the classify / send-guard decision logic are the highest-value
  surfaces. The Outlook adapter and add-in shell are excluded (they need a live
  Outlook) and validated manually.
- **Manual validation on this dev workstation:** every user-visible increment
  ships an installable add-in via the install kit; load into the dev Outlook
  and validate by hand before it moves on (`/reload-addin`).
- **Target validation:** parity-critical and COM/EDR/bitness-sensitive
  increments are re-checked on the real 32-bit target before pilot. Items that
  are explicitly **not verifiable on the dev machine** (Stormshield/S-MIME mail,
  the external banner) carry a re-verify note.

---

## Open carried-over items

Both items depend on mail this dev machine never receives (Stormshield/S-MIME,
the external-sender banner), so they are checked on the real 32-bit target and
stay out of sprint scope. The full specs and dev-side results are in the
archived version sections of [history](docs/roadmap-history.md).

- [ ] **Encrypted (Stormshield) mail: real-target verification** [#62].
      Covers v2.11.0.0 A3 (protected-mail guard and protection trace),
      v2.11.0.0 C2 (body images as removable attachments) and v2.12.0.0 A8
      (experimental re-encrypted filing). On a target workstation:
      1. **Step 0, first:** run `tools/Inspect-SmimeMsg.ps1` on a real
         Stormshield `.msg`. It must show a standard CMS envelope; record the
         key transport. RSA-OAEP keeps `ReencryptKeyTransport` = `Oaep`;
         PKCS#1 v1.5 means switching that default to `Pkcs1`; a non-standard
         envelope shelves the re-encryption feature.
      2. **Standard path, card inserted**, Diagnostic logging on: classify one
         protected mail per path (move, copy, safety copy, keep a copy) with
         "remove attachments" on. The attachments must stay, and the log must
         show `protected mail detected by <Reason>` (expected:
         `StoredMessageClass`). A `Protected mail filed by … no longer reads
         as protected` warning means Outlook filed it decrypted (a known
         object-model limit). Read the class and flags from the
         `Protection trace` lines and decide whether `PR_SECURITY_FLAGS` or
         another class prefix should drive the check (today: class-based,
         flags logged only).
      3. **Card removed:** classify still files the mail; if nothing happens,
         send the `GetSelectedItems` warning line from the log.
      4. **Body images option on** (v2.11.0.0 C2): a protected mail with body
         images keeps its images and its body untouched.
      5. **Experimental re-encryption on** (Settings → Experimental): a
         classified Stormshield mail is filed re-encrypted (stored
         `IPM.Note.SMIME`) and opens decrypted in Outlook with all its parts;
         the external banner is stripped from the rebuilt body; with the card
         removed, the mail is filed through the standard path.
      6. **Regression pass on ordinary mail:** copy, safety copy, mark as read,
         attachment removal, conversation widening.
      Close #62 once 1–6 pass, or once the feature is shelved at step 0.

- [ ] **External-sender banner not stripped on reply/forward** [#46].
      v2.9.0.0 B1 applies the strip at draft creation (`Inspector.Activate`
      rather than `NewInspector`, plus `Explorer.InlineResponse` for inline
      replies), but it is **confirmed still broken on the real target
      (2026-07-28)**: detection works (Settings → "Diagnose selected mail"
      reports a match, and the filed-copy strip works), the reply/forward
      strip does not. v2.10.0.0 D1 added Settings → "Save diagnostic file",
      which writes the body with the matched banner highlighted to
      `%LocalAppData%\RBLclass\diagnostics\`. Next steps: on the target, check
      that action (file written, match highlighted, folder created on first
      use), get a diagnostic file for a reply/forward draft, root-cause why
      the strip fails there, then fix and re-verify on the target.

---

## Phase 8 — Continuous maintenance (ongoing)

- [ ] Regression test plan executed on every Semi-Annual Enterprise channel
      update (typically January and July). Matrix: install, folder indexing,
      folder search, open-folder, classify (single + multi-destination),
      conversation widening, sub-folder creation, attachment removal, send-time
      guards, sent-item triage, settings persistence — mirrors the parity
      checklist in [docs/legacy-overview.md](docs/legacy-overview.md) §9.
- [ ] Watch the Microsoft 365 roadmap / deprecation notices for COM add-ins
      and classic Outlook.
- [ ] Annual review of the migration plan toward "new Outlook" if/when that
      becomes a corporate direction.

## Phase 9 — Security hardening (audit 2026-06-07)

A full audit found **no critical or high-severity vulnerabilities** (SQL fully
parameterized, no path traversal, no hardcoded secrets, no crypto, no outbound
network calls, COM lifetime managed via `ComRef<T>`). The items below are
**defense-in-depth and enterprise due-diligence** for a corporate security
review — several will be dropped on risk/benefit for a per-user, no-network,
single-machine add-in, but are recorded so the review trail is explicit.
Since the audit, the posture gained an opt-in update check (v2.7.0.0) and
in-box S/MIME encryption for the experimental re-encrypted filing
(v2.12.0.0); `docs/security.md` is the current reference. Completed items
are archived in [history](docs/roadmap-history.md).

### Evidence & process (highest value for sign-off)
- [ ] **Authenticode-sign the DLL + MSI** with the internal-PKI Code Signing
      cert (OID 1.3.6.1.5.5.7.3.3), timestamped — promoting today's optional
      signing to the GA default.
- [ ] **Dependency vulnerability scanning** in CI
      (`dotnet list package --vulnerable --include-transitive` + Dependabot over
      `SQLitePCLRaw`, `Microsoft.Data.Sqlite`, `Serilog`). Pin and re-verify on
      every bump — the COM host has no binding redirects
      ([[com-addin-no-binding-redirects]]).
- [ ] **SAST in CI** (Roslyn security analyzers / CodeQL, security rules as
      build errors), added to the existing `.github/workflows/ci.yml`, which
      only runs the Core tests today.
- [ ] **SBOM** per release artifact.

### Error handling & data leakage (actionable code changes)
- [ ] **Stop showing raw exception text to users.** `ShowError`
      (`RblClassAddIn.cs`) calls `MessageBox.Show(ex.ToString())`. Show a
      friendly message instead; the full exception is already logged
      (v2.7.1.0 D3).
- [ ] **Replace silent `catch { }` blocks** in `OutlookMailStore.cs` (16 left
      on 2026-11-21) with logged catches (notably the attachment-strip swallow —
      a filed copy could silently retain attachments the user believes were
      removed, [[classify-attachments-on-copy]]; surface that failure to the
      user).
- [ ] **Decide & document send-guard failure behavior.** `Application_ItemSend`
      deliberately **fails open** (CLAUDE.md forbids exceptions escaping into
      Outlook). A security team may want the external-recipient guard to **fail
      safe**. Narrow the catch so a guard *evaluation* error still prompts while
      genuine COM faults stay contained; document the stance.
### Policy & tamper resistance
- [ ] **Admin-lockable security settings.** The security-relevant toggles
      (`SendExternalWarning`, `InternalDomains`, `ForgottenAttachmentKeywords`)
      live in the user-writable SQLite `Settings` table. Add an optional **HKLM
      policy override** the security team can push by GPO (admin policy wins).
      *Likely the highest-value enterprise-specific item.*

### Input validation (low risk)
- [ ] **Cap free-text inputs:** length-limit the folder-search query and the
      settings text editors; add an **upper** bound to `MaxResults`
      (`Settings.ParseMaxResults` clamps only the lower bound today).
- [ ] **Validate `InternalDomains` format** (reject entries without a
      `domain.tld` shape).

### Minor / best-practice
- [ ] **Use `SqliteConnectionStringBuilder`** instead of `"Data Source=" +
      _dbPath`.
- [ ] **Drop file/DB paths from user-facing dialogs** unless in an explicit
      diagnostics mode.

### Accepted by design (recorded so they are not re-flagged)
- **Plaintext SQLite settings/folder index** — contents are non-secret;
  confidentiality relies on the per-user `%LocalAppData%` ACL (the correct OS
  trust boundary). No app-level encryption planned.
- **External-recipient and diagnostics dialogs** that display addresses/paths
  are **intentional UX**, not information-disclosure leaks.
