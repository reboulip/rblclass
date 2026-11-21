# Link-click phishing warning: interception study (v2.13.0.0 D1)

Internal study, not published on the website. It answers ROADMAP item **D1**
and is the input for **D3** ("warn when a link from a banner mail is opened",
GitHub issue [#37](https://github.com/reboulip/rblclass/issues/37)).

- **Question:** can RBLclass warn the user when they open a link from a mail
  that carries the external-sender banner? The preferred answer is **before**
  the link opens, with a way to cancel. Otherwise **immediately after**.
  Otherwise the feature is **not feasible**.
- **Machine:** the dev machine.
  - Classic Outlook for Microsoft 365, **64-bit, Current Channel
    16.0.20527.20196**.
  - Windows 11 ARM64 (Outlook runs as emulated x64), display at 150 %.
  - Default browser Microsoft Edge.

  The deployment target was **not** available: 32-bit Outlook, Semi-Annual
  Enterprise Channel, Stormshield and the corporate EDR. Section 6 lists what
  must be re-checked there.
- **Method:** throwaway prototypes, all removed after the study (Appendix A).
  - **In-process candidates** ran in a temporary build of the add-in loaded
    into Outlook.
  - **Observers** ran out of process as PowerShell/C# scripts (UI Automation,
    WinEvents, WMI).
  - **Input** was always real: `SetCursorPos` + `mouse_event` for clicks,
    `keybd_event` for keys, `InjectTouchInput` for touch.
  - **Fake mails:** received-looking mails in the test PST. They carried the
    learned banner, links to `https://example.com/…` (the domain reserved by
    IANA for examples) and one `mailto:` link.
- **Status of the results:** everything below was **measured** on the dev
  machine, unless it is marked *not tested* or *estimate*.
  - Timings come from millisecond-stamped logs, from one or two runs each.
  - They show orders of magnitude, not benchmarks.
  - Experiment IDs (A0, R1…, X5) refer to section 2.

## Contents

0. [Summary](#0-summary)
1. [Context and constraints](#1-context-and-constraints)
2. [Candidate-by-candidate findings](#2-candidate-by-candidate-findings)
3. [Comparison matrix](#3-comparison-matrix)
4. [Rejected options and why](#4-rejected-options-and-why)
5. [Recommended mechanism for D3](#5-recommended-mechanism-for-d3)
6. [Residual risks and target re-verification](#6-residual-risks-and-target-re-verification)
7. [Decisions needed from the user](#7-decisions-needed-from-the-user)
- [Appendix A: how the experiments were run](#appendix-a-how-the-experiments-were-run)
- [Sources](#sources)

---

## 0. Summary

1. **"Before the open" works for the usual ways of opening a link.** The full
   URL is known, and cancelling really cancels.
   - Covered gestures:
     - a **mouse click** in the reading pane or in an opened mail (plain
       click);
     - a **Ctrl+click** in a reply or forward, in its own window or inline in
       the reading pane;
     - **Enter** on a link reached with Tab, in received mail;
     - the context-menu item **"Open Hyperlink"**, with the mouse (the menu
       opened from the keyboard is *not tested*).
   - It needs three interception points. All are in-process, on Outlook's UI
     thread:
     - a **thread-scoped `WH_MOUSE` hook** for clicks;
     - a **thread-scoped `WH_GETMESSAGE` hook** for Enter;
     - a **repurposed built-in command `HyperlinkOpen`**, declared in the
       add-in's Ribbon XML, for the context menu.
   - None of them is a global hook. There is no DLL injection, no API
     patching, and no change to the stored mail.
   - **Hit test:** the mail body is a Word document. The add-in finds the link
     under the pointer through Word's own object model:
     `AccessibleObjectFromWindow(OBJID_NATIVEOM)` → `Window.RangeFromPoint` →
     `Hyperlinks`.
   - **On a banner mail,** it swallows the gesture and shows the warning. On
     **Open anyway** it calls `Hyperlink.Follow()`.
2. **Touch taps cannot be intercepted.**
   - Word takes touch input through a path that these hooks do not control.
   - Two attempts (P1, P2) neutralised every `WM_POINTER*` message of the
     contact. The link still opened.
   - For touch, the best available is a warning **immediately after** the tap,
     with the URL known (decision D-2).
3. **Rejected candidates:**
   - UI Automation events: no event at all on a real click.
   - UI Automation hit test inside Outlook: works only through a re-entrant
     pumping wait, and takes 220 ms.
   - Watching which window takes the foreground: after the fact, URL unknown,
     false positives.
   - Watching process creation: missed the launcher.
   - Rewriting links in the stored body: rejected a priori.
   - Global hooks, API patching, taking over the default-browser association:
     forbidden or out of scope.

   Word's own events fire too late to cancel a click. `WindowBeforeRightClick`
   can still suppress the context menu on a link.
4. **Measured cost:**
   - Hit test: 0–4 ms in-process (1–14 ms across processes).
   - Banner check on the rendered text: 0–10 ms.
   - The warning appears 80–100 ms after the click.
   - **No hang** in about 40 link experiments across 6 Outlook restarts.

   Links on images needed a refined hit test. I1 found the gap, and the
   refined version passed every probe point of H1 (section 2.4.1).
5. **The EDR reaction is the main unknown.**
   - A thread-scoped hook without a DLL is not injection.
   - Some EDR products still log every `SetWindowsHookEx` call.
   - *Estimate:* low to medium risk. It must be checked on the target before
     D3 ships (section 6).
6. **Side findings that matter for D2 and D3:**
   - **Existing defect: opening a banner mail in its own window strips the
     banner from that received mail and saves it.** This happens when *Strip
     banner from replies* is on.
     - Cause: `OnNewInspector` → `StripBannerFromDraft` never checks that the
       item is a draft.
     - Measured: the stored body went from 1283 to 514 characters, on a mail
       with `Sent=True`, modified at the moment it was opened.
     - Impact: received mail is modified, and D2/D3 stop seeing the banner in
       opened mails.
     - Fix it first (decision D-6, section 2.12).
   - **The HTML banner match breaks with some character sets.**
     - The same banner, stored in a mail whose character set is ISO-8859-1,
       has `’` written as `&#8217;`. `ContainsBanner` then returns false.
     - A comparison on the rendered text, letters and digits only, matched in
       every case (decision D-3).
   - **Outlook never called `OnBeginShutdown` on a normal close** of this
     profile (fast shutdown). The logs hold 38 add-in starts and 0 "shutting
     down" lines.

**Recommendation, in the requested order of preference:**

- Implement D3 **before the open** for mouse, Enter and the context menu,
  with the three mechanisms of point 1.
- For touch, warn **immediately after**, or accept the gap.
- Make both conditional on an EDR check on the 32-bit target.

Section 5 is the implementation specification. Section 7 lists the decisions.

---

## 1. Context and constraints

### 1.1 The goal (ROADMAP section D)

The decisions already taken for section D:

- **Systematic warning.** The user is warned systematically, with no action
  of their own, when they open a link from a mail that carries the
  organisation's external-sender banner.
- **Before is preferred.** A warning before the link opens is preferred; one
  immediately after is acceptable.
- **Banner test.** A mail "carries the banner" when
  `ExternalBannerStripper.ContainsBanner(htmlBody, signature)` is true.
  The signature is `SettingsKeys.ExternalBannerSignature`.
- **No warning** when no signature has been learned.

Constraints:

- **Body reads:** the body is read only when the event fires. Never on a
  selection change: that costs performance, and reading a Stormshield body
  prompts for the smart card.
- **Forbidden:** no global keyboard or mouse hook, no API patching, no
  modification of the stored mail.
- **CLAUDE.md:**
  - all Outlook object-model calls stay on the UI thread;
  - the add-in never blocks that thread and never lets an exception escape
    into Outlook;
  - no hand-written `[ComImport]` declarations of Office interfaces;
  - no new dependency without asking;
  - the same AnyCPU binary must run in 64-bit and 32-bit Outlook.

### 1.2 What Outlook provides

- **No hyperlink event** exists in the Outlook object model, in the reading
  pane or in an opened mail. D2's item events cover attachments only.
- **The mail body is rendered by Word.**
  - In every surface, the body is a `_WwB` window holding a `_WwG` window
    titled "Message".
  - The surfaces are the reading pane, an opened mail, a reply in its own
    window, and an inline reply.
  - **All of them run on Outlook's single UI thread.** Every session showed
    the same thread id for the explorer, the inspectors and their Word
    windows.
- **Access to the Word document differs by surface.**
  - The object model gives it for opened mails (`Inspector.WordEditor`) and
    for inline replies (`Explorer.ActiveInlineResponseWordEditor`).
  - It gives nothing for the **reading pane**.
  - Outlook-hosted Word returns `null` for `Document.ActiveWindow`.

### 1.3 How a link opens (measured)

| Surface | Word document state | What opens a link | Click → browser in front |
|---|---|---|---|
| Reading pane | read-only (`ProtectionType = 3`, `wdAllowOnlyReading`) | single left click, acting on the **button-down** (R5b); Enter after Tab (K1); context menu "Open Hyperlink" (R10b); touch tap (T1) | 163–405 ms |
| Opened mail | read-only | same as the reading pane (Tab + Enter: R11) | 179–365 ms |
| Reply or forward in its own window | editable | **Ctrl+click** (a plain click only moves the caret, R14a); context menu | 169 ms |
| Inline reply | editable | Ctrl+click; context menu | 188 ms |

Edge opened the link in a new tab of the window that was already open, and
came to the front. At that moment its title is still "Untitled". The page
title arrives later.

### 1.4 Criteria

Each candidate is rated on:

- before or after the open;
- whether it can cancel;
- whether the URL is known;
- coverage: reading pane, opened mail, reply/inline, Enter, context menu,
  touch;
- EDR risk on the target;
- S/MIME impact;
- performance;
- fail-open behaviour.

---

## 2. Candidate-by-candidate findings

### 2.1 Test rig (summary; details in Appendix A)

- **Fake mails** were made with the PostItem trick, then converted to
  `IPM.Note`. Each HTML body held:
  - the learned banner;
  - a link whose text differs from its target;
  - a link whose text is its URL;
  - a `mailto:` link.

  Two more fakes:
  - a control mail without banner;
  - a mail with an image link and a table "button" link.

  The banner fake had to be **stored as UTF-8** (section 2.11).
- **Every click went through a guard.** The click helper brings the target
  window to the front, then **aborts unless the window under the point belongs
  to it**. It was added after an early click hit a browser window that covered
  Outlook. Later, it correctly aborted one click while the Alt+Tab switcher
  covered the screen.
- **An observer ran outside Outlook.** It logged:
  - UI Automation Invoke, focus and text-selection events;
  - out-of-context WinEvents: foreground changes for every process, and
    every event raised by `OUTLOOK.EXE`;
  - WMI process creation, polled every 100 ms.
- **The in-process spike** was one class, created in `OnStartupComplete`.
  - It installed the hooks.
  - It re-read a small ini file at each click. That file selected the
    behaviour:
    - log only, or swallow the down, the up, or both;
    - hit test through Word or through UI Automation;
    - banner check on the HTML or on the text;
    - deferred modal, then `Follow()` or `ShellExecute`.

### 2.2 Facts shared by all candidates

- **W1. Word's object model is reachable from the window handle.**
  - `AccessibleObjectFromWindow(hwnd, OBJID_NATIVEOM, IID_IDispatch)`, called
    on a `_WwG`, returns Word's `Window`.
  - Its `Document` is the displayed mail. For an opened mail it is the same
    document as `Inspector.WordEditor`.
  - It worked in all four surfaces, **including the reading pane**.
- **The Word `Window` maps screen points to links.**
  - `Window.GetPoint` returns the screen rectangle of a range, in physical
    pixels.
  - `Window.RangeFromPoint(x, y)` returns the range under a screen point.
  - `Range.Hyperlinks(1)` exposes `Address`, `SubAddress` and
    `TextToDisplay`.
  - Cost: 1.4–14 ms from another process, 0–4 ms in-process.
- **Document state:**
  - read-only mail documents report `ProtectionType = 3`;
  - compose documents are editable.
- **Word follows a link on the button-down** in read-only documents.
  - Swallowing only the up did not stop the link (R5b).
  - Swallowing only the down did (R6).
- **One physical click reaches a `WH_MOUSE` hook several times** (R1):
  - the down arrives as 2 `HC_NOREMOVE` peeks, then 1 `HC_ACTION`;
  - the up arrives as up to 8 peeks, then 1 action.

### 2.3 Candidate 1: UI Automation (events, or `FromPoint` at click time)

- **U1 (outside Outlook).** UI Automation exposes each link as a `Hyperlink`
  element.
  - Patterns: Invoke, Value, Text, ScrollItem. `Value` holds the URL.
  - `FindAll` takes 54 ms and `FromPoint` 13 ms.
- **A0 (outside Outlook, opened mail, real click on a link).** UI Automation
  raised **no `InvokePattern.InvokedEvent`, no focus event and no
  text-selection event**.
  - Only WinEvents fired, from the `_WwG`, about 220 ms after the click:
    `EVENT_OBJECT_TEXTSELECTIONCHANGED` (0x8014) and four events in the
    OEM-defined range (0x010A, 0x010C, 0x011B, 0x0121).
  - Edge came to the front 365 ms after the click.
  - A UI Automation client driven by events therefore never learns about the
    click in time. These WinEvents come after the fact, are undocumented and
    cannot cancel.
- **R9 (inside Outlook).** The `WH_MOUSE` hook ran `AutomationElement.FromPoint`
  on a background MTA thread, with the UI thread waiting at most 300 ms.
  - It returned the right `Hyperlink` and URL in **220 ms**.
  - It worked only because the CLR's wait on an STA thread pumps messages.
    Word's UI Automation provider therefore ran **re-entrantly inside the
    hook**.
  - Microsoft's UI Automation threading guidance warns against client calls
    that depend on the provider's UI thread, because of the deadlock risk.
  - The Word object model gives the same answer in at most 4 ms, without
    re-entrancy.

**Verdict:** not usable as a trigger. As a hit test, the Word object model
beats it on every count.

### 2.4 Candidate 2: thread-scoped `WH_MOUSE` hook plus hit test

Installed with
`SetWindowsHookEx(WH_MOUSE, proc, IntPtr.Zero, GetCurrentThreadId())` in
`OnStartupComplete`, which runs on the UI thread. It installed successfully
(`GetLastError` 0) at every one of the five restarts with the hook.

| ID | Surface | Spike setting | Observed |
|---|---|---|---|
| R1 | reading pane | log only | The hook sees the down and the up. Hit test: 4 ms on first use, then 0–1 ms. `WindowSelectionChange` fires 120 ms after the up, and Edge comes to the front 36 ms later. |
| R2, R4 | reading pane | swallow the down and the up that follows | **Nothing opens.** Decision takes about 6 ms once caches are warm. Returning 1 at the first `HC_NOREMOVE` peek removes the message: no `HC_ACTION` follows. |
| R5b | reading pane | let the down through, swallow the up | **The link opens**, so Word acts on the down. |
| R6 | reading pane | swallow the down, let the up through | Nothing opens. No side effect. |
| R7 | reading pane | swallow, deferred modal, "Yes", `Hyperlink.Follow()` | The modal appears 98 ms after the click, and `Follow()` opens the link. **Windows dropped the matching up** because the modal had disabled the owner window. A "swallow the next up" flag must therefore expire. |
| R8 | reading pane | swallow, modal, "No" | Nothing opens. Modal after 82 ms. |
| R12 | reading pane, **no banner** | swallow mode | Not swallowed: the link opens normally. The hook added about 17 ms (banner checks on first use). |
| R13, R13b | opened mail | swallow; swallow without banner requirement | Hook, hit test and Word events all work. R13: the banner check found no banner because the existing strip defect had just removed it (section 2.12), so the link opened. R13b: swallowing stops the open. |
| R14a/b/c | reply in its own window | plain click; Ctrl+click swallowed; Ctrl+click logged | Plain click: caret only. Ctrl+click logged: the link opens (169 ms). Ctrl+click swallowed: nothing opens. |
| R15a/b | inline reply | Ctrl+click swallowed; Ctrl+click logged | Nothing opens, then the link opens (188 ms). `Explorer.Selection(1)` is still the original mail. |
| R10b | reading pane | context menu "Ouvrir le lien hypertexte" | **The link opens.** The click lands on the menu window (`NetUIHWND`), not on Word: not covered by this hook. |
| R11 | opened mail | Tab, Tab, Enter | **The link opens.** Tab moves the selection from link to link, and Enter follows it: not covered by this hook. |
| T1 | reading pane | touch tap | Only `WM_POINTER*` messages reach Word, no `WM_LBUTTON*`. The hook is never called and **the link opens**. |
| I1 | image link | swallow | **Missed:** the hit test returned no link (boundary case, 2.4.1), so the link opened. |
| I2 | table "button" link, on its text | swallow | Nothing opens. |
| I3 | button padding, off the text | log | Word follows the link, and the hit test agreed. |

#### 2.4.1 Hit test refinement (H1)

`RangeFromPoint` on an **image link** returns a collapsed range exactly at
the start of the hyperlink's range. The image link spanned positions 18–116,
an inline-shape hyperlink. A collapsed range on that boundary reports
`Hyperlinks.Count = 0`. A click on the left half of the first character of a
text link behaves the same way.

The refined test runs in three steps:

1. Take `p = RangeFromPoint(x, y)`. If `p` lies inside a hyperlink, use it.
2. Otherwise, take the one-character range `r1 = (p, p+1)`. If `r1` holds a
   hyperlink, accept it only when the point lies inside **the screen box of
   the whole hyperlink range**, from `GetPoint(hl.Range)`, with ±2 px
   tolerance.
3. The box of `r1` itself is useless. A hyperlink range starts with its
   hidden field code, so `GetPoint(r1)` returned a width of 0.

H1 probed 11 points, outside Outlook:

| Probe point | Result |
|---|---|
| image centre, corner and left edge | link (correct) |
| left half of a text link's first character | link (correct) |
| 12 px left of a text link | none (correct) |
| right of a text link's end | none (correct) |
| line margin | none (correct) |
| button text and button padding | link, as Word itself does (I3) |
| plain text | none (correct) |

#### 2.4.2 Item, sender and banner at click time

- **Which mail was clicked.** The spike mapped the clicked `_WwG` to its
  owner: it compared `GetAncestor(hwnd, GA_ROOT)` with
  `IOleWindow.GetWindow()` of the active explorer and of each inspector.
  - It identified the explorer and the inspector correctly every time.
  - Inline reply: `Explorer.ActiveInlineResponse` is the draft, and
    `Explorer.Selection(1)` is the original mail.
- **Sender:** `SenderName` and the sender address were readable from the
  same item.
- **Banner check:** 0–10 ms, see section 2.11.

#### 2.4.3 Assessment

- **Timing and control:** acts before the open, can cancel, knows the URL.
- **Coverage:**
  - covered: reading pane, opened mail, Ctrl+click in a reply in its own
    window and in an inline reply;
  - not covered: Enter, the context menu, touch.
- **EDR** (*estimate*): low to medium.
  - Thread-scoped hook with no DLL, so no injection and no cross-process
    effect.
  - The API call itself may still be logged.
- **S/MIME:** no impact. The stored mail is untouched.
- **Cost:** a mouse move costs one integer comparison plus `CallNextHookEx`
  (not instrumented). A click on a Word window costs a 0–4 ms hit test, plus
  the banner check when the click lands on a link.
- **Fail-open:** every error passes the click through.

### 2.5 Candidate 2b: thread-scoped `WH_GETMESSAGE` hook

The documented `GetMsgProc` contract lets the hook modify the retrieved
message. The spike neutralised a message by writing `WM_NULL` into
`MSG.message`.

- **G1.** Turning `WM_LBUTTONDOWN` into `WM_NULL` at `PM_REMOVE` cancels a
  click in the reading pane, so it is an alternative to `WH_MOUSE`.
- **K1.** The Enter key can be cancelled the same way.
  - Condition: `WM_KEYDOWN` with `VK_RETURN` on a `_WwG`, while Word's
    `Selection` is inside a hyperlink of a read-only document
    (`ProtectionType = 3`).
  - The hook turns the key into `WM_NULL`: **no open**.
  - The URL comes from `Selection.Hyperlinks(1)`.
- **P1.** All messages of the touch contact were turned into `WM_NULL` at
  `PM_REMOVE`: `WM_POINTERDOWN`, then `UPDATE`, `UP` and `LEAVE`. **The link
  still opened.**
- **P2.** The same was done from the first in-contact `WM_POINTERENTER`, at
  **both** `PM_NOREMOVE` and `PM_REMOVE`. **The link still opened**, about
  530 ms after the tap.
  - Word therefore handles touch through a path that does not depend on the
    messages as the thread retrieves them. Pointer-frame or
    interaction-context APIs are likely; this was not investigated further.
  - The hook does see where the finger went down, so it can hit-test that
    point.

**Assessment:**

- Mouse: like candidate 2, verified in the reading pane only.
- **Enter: covered.**
- Touch: not interceptable, only observable.

### 2.6 Candidate 3: out-of-context `SetWinEventHook(EVENT_SYSTEM_FOREGROUND)`

Every run logged foreground changes, with no DLL in Outlook (out of context).

- **Too late, and the URL stays unknown.** The browser came to the front
  160–410 ms after the click. Its title at that moment was "Untitled": the
  page title arrived later as a name change.
- **Noisy.** Foreground events also come with no click at all:
  - Windows shell windows such as `ForegroundStaging` and the Alt+Tab
    switcher (`XamlExplorerHostIslandWindow`);
  - the user switching to the browser by hand, which happened several times
    during the study.
- **Needs correlation.** It covers every gesture, touch included. But it is
  usable only alongside another signal saying "a link on a banner mail was
  just activated".

**Assessment:** a fallback only, and the weakest one.

### 2.7 Candidate 4: rewriting links in the stored body

Not prototyped. It was rejected a priori; the reasons are in section 4.

### 2.8 Candidate 5: Word events

- **Outlook-hosted Word does raise Word application events.** The spike
  subscribed through `Window.Application` from W1.
  - `WindowSelectionChange` fires on clicks in all surfaces, on Tab between
    links, and on touch taps.
  - `WindowBeforeRightClick` fires on a right-click.
- **`WindowSelectionChange` comes too late.** It fires **after the button-up**,
  17–275 ms after it across the runs, and only 28–62 ms before the browser
  comes to the front. It has no `Cancel`.
- **`WindowBeforeRightClick` has a `Cancel`.** Setting it on a link
  **suppressed the context menu**, as verified. That is a blunt alternative to
  candidate 6: it also removes "Copy Hyperlink".
- **Cost of subscribing:** event sinks need Word's interop assembly or a
  hand-built `IDispatch` sink. The recommendation does not use them.

### 2.9 Added candidate 6: repurposing the built-in `HyperlinkOpen` command

The spike added
`<commands><command idMso="HyperlinkOpen" onAction="…"/></commands>` to the
add-in's Ribbon XML. The study changed only the French variant, the one
loaded here.

- **X1, X2: the wrong signature breaks the menu item.** With the
  push-button callback signature, `(IRibbonControl control, ref bool
  cancelDefault)`:
  - the callback was never reached;
  - the menu item **did nothing at all**, with or without cancel: the link
    could no longer be opened from the menu.

  UI Automation reports the item as a check-menu item
  (`NetUITWBtnCheckMenuItem`).
- **X3: the toggle signature works.** With `(IRibbonControl control, bool
  pressed, ref bool cancelDefault)`:
  - the callback fired with `control.Id = "HyperlinkOpen"` and
    `control.Context` set to the `Explorer`;
  - the URL came from Word's `Selection`, because the right-click had put the
    selection on the link;
  - setting `cancelDefault = true` stopped the open.
- **X4.** Same callback, returning `false`: the default ran and the link
  opened.
- **X5.** In an opened mail, the callback fired with `Context` set to the
  `Inspector`, and cancelling worked.
- **Custom UI still valid.** The RBLclass ribbon tab stayed present, so Office
  accepted the custom UI.

**Assessment:**

- Official extensibility: no EDR concern.
- Acts before the open, can cancel, knows the URL.
- *Not tested:* the context menu opened from the keyboard (Shift+F10) and the
  menu in a reply window. It is the same command, so both are expected to work.
- Risk: a callback with the wrong signature fails **closed**.

### 2.10 Added candidate 7: watching process creation

- WMI `__InstanceCreationEvent WITHIN 0.1` needs no admin rights. In every run
  it saw only an Edge **renderer** process, whose parent was the Edge already
  running.
- The short-lived launcher, the process that would carry the URL on its
  command line, was never caught.
- Kernel process tracing (ETW) needs admin rights. The approach is after the
  fact anyway.

**Verdict:** rejected.

### 2.11 Banner detection at click time

- **HTML, as D2 defines it.** Find the displayed item (`Selection(1)`, the
  inspector, or the inline reply), then run
  `ContainsBanner(HTMLBody, signature)`. It took 1–9 ms; the 40 KB body of a
  reply took 6 ms.
- **The HTML match depends on the stored character set (measured).**
  - The first banner fake was stored as ISO-8859-1 (code page 28591). Outlook
    wrote U+2019 `’` as `&#8217;`.
  - `ContainsBanner` then failed, in both its exact and its
    whitespace-tolerant forms, although the banner looked identical on screen.
  - The same fake stored with code page 65001 matched.
  - Real banner mails will fail the same way when the sender's encoding cannot
    represent a character of the banner (*frequency unknown*, re-check on the
    target).
- **Rendered text.** Read the text of the `_WwG`'s `Document.Content`. This
  involves no Outlook item and no body read.
  - Reduce both sides to letters and digits, in lower case, then test for a
    substring.
  - It matched both banner fakes and rejected the control mail.
  - Cost: 10 ms on first access, 0–2 ms afterwards.
- **S/MIME.** The rendered text is the decrypted content already on screen,
  and reading it does not touch the Outlook item. By construction it cannot
  trigger a smart-card prompt. The HTML path on an encrypted item was *not
  tested*: it carries the Stormshield prompt risk.

### 2.12 Side findings

- **Existing defect: received mail loses its banner when opened.**
  - **Code:** `OnNewInspector` (`RblClassAddIn.cs:754`) subscribes to
    `Inspector.Activate` for **every** inspector. Its handler calls
    `_mailStore.StripBannerFromDraft(inspector.CurrentItem, sig)`
    (`RblClassAddIn.cs:783`).
    `StripBannerFromDraft` (`OutlookMailStore.cs:1816`) calls
    `StripBannerOnLiveMail` with `isDraft: true`. That method rewrites
    `HTMLBody` and calls `Save()` (`OutlookMailStore.cs:1848-1851`) **without
    checking `mail.Sent`**.
  - **Effect:** with *Strip banner from replies* on, double-clicking a
    received banner mail removes the banner **from the received mail**, for
    good.
  - **Evidence (D1-BANNER2):** the add-in logged "Stripped the
    external-sender banner from mail body …" at the moment it was opened. The
    stored body went from 1283 to 514 characters, with `Sent=True` and the
    modification time of the open.
  - **Impact on D2/D3:** after this, warnings can never fire for that mail,
    in its window or later in the reading pane.
  - **Fix:** strip only drafts (`mail.Sent == false`). It is a separate fix
    item (D-6).
- **Outlook's fast shutdown skips `OnBeginShutdown`.** It never runs on a
  normal close of this profile: 38 add-in starts, 0 "shutting down" log lines.
  - Thread hooks die with the process, so this is harmless for D3 at exit.
  - `OnDisconnection` does run when the user disables the add-in while Outlook
    is open (Office extensibility contract, *not tested here*). D3 must unhook
    there, or Outlook would call a released delegate.
- **Hook mechanics to respect** (sections 2.2, 2.4):
  - returning 1 from an `HC_NOREMOVE` call removes the message;
  - the matching up can be dropped by Windows when a modal opens;
  - `WH_MOUSE` never sees touch.
- **Test setup note (not product behaviour).** The dev profile had "Always
  preview messages" switched off in the reading-pane options, so the reading
  pane showed nothing on selection. It was switched on for the tests, and the
  registry value it creates (`ReadingPaneSelectItem`) was removed afterwards.

---

## 3. Comparison matrix

Legend for the matrix:

- "expected": not tested, but the same code path as a tested case.
- **EDR risk** is an *estimate*: none of it could be verified on the target.

| Candidate | When | Cancel | URL | Reading pane | Opened mail | Reply/inline (Ctrl+click) | Enter | Context menu | Touch | EDR risk | S/MIME impact | Cost | Fail-open | Evidence |
|---|---|---|---|---|---|---|---|---|---|---|---|---|---|---|
| 1a. UIA events | never | no | no | no event | no event | not tested | not tested | not tested | not tested | low | none | n/a | n/a | A0 |
| 1b. UIA `FromPoint` in-process | before (needs a hook to trigger it) | through the hook | yes | yes | expected | expected | no | no | no | low, plus the hook's | none | 220 ms on first use, re-entrant | risky (pumping wait) | R9, U1 |
| **2. `WH_MOUSE` + Word hit test** | **before** | **yes** | **yes** | **yes** | **yes** | **yes** | no | no | no | low to medium | none | 0–4 ms, plus banner 0–10 ms | yes | R1–R15, I1–I3, H1 |
| **2b. `WH_GETMESSAGE`** | **before** | **yes** | **yes** | yes (G1) | expected | expected | **yes (K1)** | no | **no (P1, P2)** | low to medium | none | as candidate 2 | yes | G1, K1, P1, P2 |
| 3. Foreground WinEvent | after (+160–410 ms) | no | no (title "Untitled") | with correlation | with correlation | with correlation | with correlation | with correlation | with correlation | low | none | negligible | n/a | all runs |
| 4. Link rewriting | before | yes | yes | yes | yes | yes | yes | yes | yes | high | breaks signatures; impossible on encrypted bodies | n/a | n/a | not prototyped |
| 5. Word events | after the button-up, before the browser | no (right-click: menu only) | yes | yes | yes | yes | sees Tab | can suppress the menu | sees the selection | low | none | negligible | yes | R1, R3, R11, T1 |
| **6. Repurpose `HyperlinkOpen`** | **before** | **yes** | **yes** (selection) | n/a | n/a | menu, expected | n/a | **yes** (mouse; keyboard expected) | n/a | none | none | negligible | only with the exact toggle signature, otherwise fail-closed | X1–X5 |
| 7. WMI process watch | after | no | no (launcher missed) | n/a | n/a | n/a | n/a | n/a | n/a | medium | none | polling | n/a | all runs |

**Recommended combination:** candidates 2 + 2b (Enter only) + 6.

- Before the open, with cancel and the URL: mouse, Ctrl+click, Enter and
  context menu, in all four surfaces.
- Touch: a warning immediately after (optional, D-2).

---

## 4. Rejected options and why

**Rewriting links in the stored body** (to a local handler through a custom
HKCU protocol) was **rejected a priori**. It cannot be fixed by
implementation care:

- It **modifies received mail.** That is forbidden by section D, and it
  destroys the mail's value as evidence when a phishing attempt is reported.
- It **breaks S/MIME signatures**: the signed body no longer matches its
  signature. It is **impossible on encrypted mail**, whose body cannot be
  rewritten without decrypting it and re-encrypting it to the user.
- It **spreads the rewrite.** Rewritten copies sync back to the server (IMAP or
  Exchange) and reach every client and device of the user. If the add-in is
  removed, the links stay broken.
- It **adds an executable and a trust boundary.** A protocol handler is a new
  binary that receives attacker-controlled URLs from the command line. That is
  a classic attack surface, and EDR products watch it.
- It **interferes with gateway or Safe Links rewriting.** The mail's displayed
  text and targets diverge even more.

Other rejected options:

- **UI Automation events** (1a): no event on a real click (A0).
- **UI Automation hit test inside Outlook** (1b): it needs a trigger anyway.
  It takes 220 ms and runs Word's provider re-entrantly inside a hook, while
  the Word object model gives the same answer in at most 4 ms.
- **An out-of-process UI Automation or WinEvent observer:** a second
  executable, a new trust boundary, and still no cancel.
- **Global hooks** (`WH_MOUSE_LL`, `WH_KEYBOARD_LL`, or a global
  `WH_MOUSE`): forbidden by section D, and they match the keylogger pattern
  EDR products look for.
- **API patching** (detours on `ShellExecuteExW` or `HlinkNavigate`):
  forbidden by section D, and highly suspicious to EDR products.
- **Process watching** (7): after the fact, and it missed the launcher. The
  reliable form (ETW) needs admin rights.
- **Taking over the default-browser or http/https association:**
  - it changes behaviour for every application;
  - since Windows 8 the user choice is protected by a hash and cannot be set
    silently;
  - it needs an executable.
- **Word events alone** (5): they fire too late to cancel a click. They stay
  useful as signals.
- **Suppressing the whole context menu on links** (`WindowBeforeRightClick`
  with `Cancel`): it works, but also removes "Copy Hyperlink", which users
  need to inspect a link. It remains a fallback if candidate 6 fails on the
  target.
- **Foreground watching alone** (3): after the fact, no URL, false positives.
- **`WH_GETMESSAGE` for touch** (2b): two attempts failed (P1, P2).
- **Reading mail as plain text** (an Outlook option): hostile to users, and
  URLs are still clickable.
- **Organisational controls,** such as Defender for Office 365 Safe Links or
  gateway URL rewriting with time-of-click checks. These are the best answer
  to link phishing, but they sit outside the add-in, and the project takes no
  dependency on Exchange Online. They are worth mentioning to IT, not
  something RBLclass can build.
- **Subclassing the `_WwG` windows** (`SetWindowSubclass`): *not prototyped*.
  - For posted mouse and key messages it would cancel exactly as
    `WH_GETMESSAGE` does, without any `SetWindowsHookEx` call.
  - It would have the same limit on touch.
  - It needs window tracking: a new `_WwG` appears per inspector and per
    inline reply, and the reading pane's window is recreated.
  - It is kept as the **EDR fallback** if the target flags thread hooks.

---

## 5. Recommended mechanism for D3

This section is written so that D3 can be implemented from it alone. It uses
the ROADMAP names: setting `PhishingWarnLinks` and D2's
`PhishingWarningPolicy`.

### 5.1 Overview

```
Outlook UI thread (STA), all in-process, nothing global
├─ WH_MOUSE thread hook ───── left button down/dblclk on a "_WwG" window
│                              → hit test → banner? → swallow + post warning
├─ WH_GETMESSAGE thread hook ─ WM_KEYDOWN VK_RETURN on a read-only "_WwG"
│                              with the selection in a hyperlink → WM_NULL + post warning
├─ Ribbon <command idMso="HyperlinkOpen"> ─ context menu "Open Hyperlink"
│                              → banner? → cancelDefault = true + post warning
└─ Warning (WPF modal, posted with Dispatcher.BeginInvoke after the hook returns)
     Open anyway → Hyperlink.Follow()   Cancel → nothing
```

The **Word object model is reached by late binding only.**

- The path is `AccessibleObjectFromWindow(OBJID_NATIVEOM)` → `IDispatch`,
  then `Type.InvokeMember` calls.
- No Word interop assembly is referenced, and no Word interface is declared.
  This follows CLAUDE.md, and C1 makes the same choice.
- The only `[ComImport]` is `IOleWindow` (`00000114-0000-0000-C000-000000000046`).
  It is a standard OLE interface, not an Office one. Declare its two methods
  with `[PreserveSig]`.

### 5.2 Core (`RBLclass.Core`, netstandard2.0, unit-tested)

1. **`ExternalBannerStripper.ContainsBannerText(string plainText, string
   bannerSignature)`** (new).
   - Build the signature text:
     1. strip the tags with the regex `<[^>]+>` → "";
     2. HTML-decode with `WebUtility.HtmlDecode`;
     3. keep only `char.IsLetterOrDigit`, in lower-invariant case.
   - Reduce `plainText` the same way, then return whether it contains the
     signature text, by ordinal comparison.
   - Return false when either side is null or empty, or when the signature
     reduces to fewer than 16 characters (guard against trivial
     signatures).
   - Tests:
     - the same banner with `’`, `'` or `&#8217;` matches;
     - `<b>EXTERNAL SENDER</b>` matches Word text with `\v` and `\r` breaks;
     - a missing last word does not match;
     - an empty or whitespace signature returns false;
     - a null text returns false.
2. **`PhishingLinkFilter`** (new, static).
   - `IsWarnable(string url)` returns false for null or whitespace, for `#…`
     (an in-document anchor) and for `mailto:` (decision D-4). It returns
     true for everything else: http, https, ftp, file, and unknown schemes.
   - `HostForLog(string url)` returns `scheme://host`, or `"(unparsed)"`.
   - Tests for each case.
3. **`PhishingWarningPolicy`** (created by D2; D3 adds one method).
   - `ShouldWarnLink(bool settingOn, string signature, string renderedText,
     string url)` returns:
     `settingOn && !IsNullOrWhiteSpace(signature) && IsWarnable(url) &&
     ContainsBannerText(renderedText, signature)`.
   - Decision D-3 may change the banner source to the HTML body.
   - Truth-table tests.
4. **`LinkWarningMessage`** (model): sender display name, sender SMTP address,
   URL, displayed text (`TextToDisplay`), and a flag set when the displayed
   text differs from the URL.

### 5.3 Settings and strings

- **The setting:**
  - `SettingsKeys.PhishingWarnLinks = "PhishingWarnLinks"`, a bool, default
    **true**;
  - a `Settings.PhishingWarnLinks` property, loaded and saved like the other
    flags;
  - in the Settings window, a checkbox next to the external-banner options
    and next to D2's attachment option.
- **Strings:** add these keys to `Strings.resx`, `Strings.fr.resx` and
  `Strings.de.resx`. If D2 already defines a "Possible phishing" title key,
  reuse it.

| Key | EN | FR | DE |
|---|---|---|---|
| `PhishingLink_Title` | Possible phishing | Hameçonnage possible | Möglicher Phishing-Versuch |
| `PhishingLink_Text` | This link comes from an external email. It may be a phishing attempt. | Ce lien provient d'un e-mail externe. Il peut s'agir d'une tentative d'hameçonnage. | Dieser Link stammt aus einer externen E-Mail. Es könnte sich um einen Phishing-Versuch handeln. |
| `PhishingLink_Sender` | From: {0} <{1}> | Expéditeur : {0} <{1}> | Absender: {0} <{1}> |
| `PhishingLink_Url` | Link destination: {0} | Destination du lien : {0} | Linkziel: {0} |
| `PhishingLink_DisplayedText` | Shown in the email as: {0} | Affiché dans l'e-mail comme : {0} | In der E-Mail angezeigt als: {0} |
| `PhishingLink_OpenAnyway` | Open anyway | Ouvrir quand même | Trotzdem öffnen |
| `PhishingLink_Cancel` | Cancel | Annuler | Abbrechen |
| `PhishingLink_AfterTouch_Text` (only with D-2) | You opened a link from an external email. It may be a phishing attempt. | Vous avez ouvert un lien provenant d'un e-mail externe. Il peut s'agir d'une tentative d'hameçonnage. | Sie haben einen Link aus einer externen E-Mail geöffnet. Es könnte sich um einen Phishing-Versuch handeln. |
| `Settings_PhishingWarnLinks` | Warn before opening a link from an external email | Avertir avant d'ouvrir un lien provenant d'un e-mail externe | Vor dem Öffnen eines Links aus einer externen E-Mail warnen |

### 5.4 Add-in classes (`RBLclass.AddIn`, net48)

These are UI-shell concerns: windows, messages and the hosted editor. They
belong in the add-in layer, like `NavRailInset` and `PaneFocus`.

- **`Phishing/NativeMethods.cs`.**
  - Functions: `SetWindowsHookEx`, `UnhookWindowsHookEx`, `CallNextHookEx`,
    `GetCurrentThreadId`, `GetClassName`, `GetAncestor` (`GA_ROOT = 2`),
    `GetKeyState`, `GetFocus`, and `AccessibleObjectFromWindow` from oleacc.
  - Structures: `MOUSEHOOKSTRUCT` (`POINT pt; IntPtr hwnd; uint
    wHitTestCode; IntPtr dwExtraInfo`) and `MSG` (`IntPtr hwnd; uint message;
    IntPtr wParam; IntPtr lParam; uint time; POINT pt`). Every handle-sized
    field is an `IntPtr`, so the same binary works in 32-bit and 64-bit
    Outlook.
  - `IOleWindow` (`[ComImport]`, `InterfaceIsIUnknown`, `[PreserveSig] int
    GetWindow(out IntPtr)`, `[PreserveSig] int ContextSensitiveHelp(bool)`).
- **`Phishing/WordLinkAccess.cs`.** Late-bound access to the Word document
  behind a `_WwG`:
  - `TryGetWindow(hwnd)` uses `OBJID_NATIVEOM = 0xFFFFFFF0` and
    `IID_IDispatch`;
  - `HitTest(window, x, y)` implements section 5.8;
  - `IsReadOnly(window)` tests `Document.ProtectionType == 3`;
  - `CtrlClickRequired(window)` reads `Application.Options.CtrlClickHyperlinkToOpen`;
  - `SelectionLink(window)` reads `Selection.Hyperlinks(1)`;
  - `RenderedText(window)` reads `Document.Content.Text`;
  - `Follow(…)`.

  It contains a tiny `LateBound` helper: `Get(obj, name)`, `Call(obj, name,
  args)`, and `GetPoint` through `InvokeMember` with a `ParameterModifier`
  that marks its four out-parameters as by-reference. Every COM object it
  obtains is released in `finally`, through `ComRef<T>`.
- **`Phishing/LinkMailResolver.cs`.** From the clicked `_WwG` to the Outlook
  item and its sender, as section 5.9 describes.
- **`Phishing/LinkClickGuard.cs`.**
  - Owns both hooks, with each delegate held in a field so the GC keeps it.
  - Runs the per-click state machine (5.5, 5.6) and posts the warning.
  - Implements `IDisposable`; `Dispose` is idempotent.
- **`Views/PhishingWarningWindow.xaml`.** Shared with D2, with an optional
  URL block. If D2 has not landed yet, create the window with both modes.
- **`RblClassAddIn.cs` changes:**
  - a `LinkClickGuard _linkGuard` field;
  - in `OnStartupComplete`, after settings, logging and localization:
    `try { _linkGuard = new LinkClickGuard(...); _linkGuard.Install(); }
    catch (Exception ex) { TryLog(...); }`;
  - `_linkGuard?.Dispose()` in **both** `OnBeginShutdown` and
    `OnDisconnection`;
  - the public ribbon callback described in section 5.7.
- **Ribbon files:** `Ribbon.xml`, `Ribbon.fr.xml` and `Ribbon.de.xml` each get
  the same block, placed before `<ribbon>`:

  ```xml
  <commands>
    <command idMso="HyperlinkOpen" onAction="OnHyperlinkOpenRepurposed" />
  </commands>
  ```

  Check that `ResourceParityTests` still passes.

### 5.5 Mouse path (`WH_MOUSE`, thread-scoped)

**Install** with `SetWindowsHookEx(7 /*WH_MOUSE*/, _mouseProc, IntPtr.Zero,
GetCurrentThreadId())`, called on the UI thread from `OnStartupComplete`.
If the result is `IntPtr.Zero`, log `Marshal.GetLastWin32Error()` and carry
on without the guard.

**Hook procedure**, with the whole body inside `try { … } catch { log once;
return CallNextHookEx(...); }`:

1. If `nCode < 0`, return `CallNextHookEx`.
2. Read `msg = (int)wParam`. Act only on `WM_LBUTTONDOWN` (0x201),
   `WM_LBUTTONDBLCLK` (0x203), `WM_LBUTTONUP` (0x202) and `WM_RBUTTONDOWN`
   (0x204). Return `CallNextHookEx` for everything else, and test this before
   any other work: it is the hot path for mouse moves.
3. **Up:**
   - when `_swallowUp` is set and less than 1500 ms old: clear it and return
     1;
   - when it is older: clear it and pass the up.

   The up can be dropped by Windows when the warning opens (R7), hence the
   expiry. Swallowing the up is cleanup, not cancellation: R6 showed the down
   alone decides.
4. **Right-down on a `_WwG`:** remember `_lastRightWwg = hwnd` and
   `_lastRightTick`, for section 5.7. Pass it.
5. **Down (one decision per physical click).**
   - Build `key = (hwnd, pt)`. If it equals `_pendingKey` and is less than
     2000 ms old, return the stored decision. Clear `_pendingKey` on
     `HC_ACTION` (`nCode == 0`).
   - Otherwise, store the key and decide now, even when `nCode` is
     `HC_NOREMOVE` (3). Returning 1 at a peek removes the message (R2).
6. **Decide:**
   1. The class name must be `_WwG`. A warning open or pending means: return
      1 silently (the second click of a double-click).
   2. Load the settings snapshot:
      `PhishingWarnLinks && signature != ""`. A link click is rare, so a
      `Settings.Load` at that moment is acceptable. Otherwise pass.
   3. Get `win = WordLinkAccess.TryGetWindow(hwnd)`. On failure, pass.
   4. Check that the click would follow the link:
      `follows = IsReadOnly(win) || (CtrlClickRequired(win) ?
      GetKeyState(VK_CONTROL) < 0 : true)`. If it would not, pass. A plain
      click in a reply only moves the caret.
   5. Run `link = HitTest(win, pt.X, pt.Y)` (5.8). No link, or a URL that
      `PhishingLinkFilter.IsWarnable` rejects: pass.
   6. Get `text = RenderedText(win)`. If
      `PhishingWarningPolicy.ShouldWarnLink(...)` is false, pass.
   7. Swallow:
      - store `_swallowUp = true` and `_swallowUpSince = now`;
      - store `_pendingDecision = true` and `_warningPending = true`;
      - capture `(hwnd, pt, url, displayed text)`;
      - call `Dispatcher.BeginInvoke(DispatcherPriority.Normal, ShowWarning(…))`;
      - return 1.

      Never show UI, wait, or pump messages inside the hook.

### 5.6 Enter path (`WH_GETMESSAGE`, thread-scoped)

**Install** with `SetWindowsHookEx(3 /*WH_GETMESSAGE*/, _getMsgProc,
IntPtr.Zero, GetCurrentThreadId())`.

**Hook procedure:**

1. If `nCode < 0`, or `wParam != PM_REMOVE (1)`, return `CallNextHookEx`.
2. Read `message = Marshal.ReadInt32(lParam, IntPtr.Size)`. Anything other
   than `WM_KEYDOWN` (0x100): return `CallNextHookEx` (hot path).
3. Read `wParam` from the `MSG` structure. It must be `VK_RETURN` (0x0D). The
   window must be a `_WwG`, with no Ctrl, Alt or Shift held.
4. Get `win`. Require `IsReadOnly(win)`: in a reply, Enter inserts a
   paragraph. Then `link = SelectionLink(win)`; if there is none, pass.
5. Apply the same filter and policy as the mouse path. To swallow:
   - write `WM_NULL` with `Marshal.WriteInt32(lParam, IntPtr.Size, 0)` (K1).
     The application then never calls `TranslateMessage` on Enter, so no
     `WM_CHAR` follows;
   - post the warning;
   - keep turning repeated `WM_KEYDOWN` Enter into `WM_NULL` while the
     warning is pending (auto-repeat).
6. Always return `CallNextHookEx`.

### 5.7 Context-menu path (repurposed `HyperlinkOpen`)

**The callback must be public on `RblClassAddIn`,** whose AutoDispatch class
interface is how Office finds callbacks. It must have **exactly** the toggle
signature (X3):

```csharp
public void OnHyperlinkOpenRepurposed(Office.IRibbonControl control, bool pressed, ref bool cancelDefault)
```

**A different signature leaves the menu item dead** (X1, X2). The body:

1. Start from `cancelDefault = false`, then enter `try`.
2. Find the `_WwG`:
   - `_linkGuard.LastRightClickWwg`, when it is less than 30 s old;
   - otherwise `GetFocus()` when it is a `_WwG` (the menu opened from the
     keyboard; *not tested*).
3. Read `link = SelectionLink(win)`: the right-click puts the selection on
   the link (X3).
4. Apply the same filter and policy. To warn, set `cancelDefault = true` and
   post the warning.
5. In `catch`: log, and leave `cancelDefault = false` (fail-open).

`control.Context` is the `Explorer` or the `Inspector` (X3, X5). The resolver
can use it to find the mail.

### 5.8 Hit test (validated by H1)

```
p  = win.RangeFromPoint(x, y)                       // collapsed range
if p.Hyperlinks.Count > 0: return p.Hyperlinks(1)   // strictly inside a link
r1 = win.Document.Range(p.Start, min(p.Start + 1, Document.Content.End))
if r1.Hyperlinks.Count > 0:
    hl = r1.Hyperlinks(1)
    win.GetPoint(out l, out t, out w, out h, hl.Range)    // box of the whole link
    if l-2 <= x <= l+w+2 and t-2 <= y <= t+h+2: return hl // image link / left edge
return none
url = hl.Address + (hl.SubAddress non-empty ? "#" + hl.SubAddress : "")
```

The hook's `pt` and Word's `RangeFromPoint` use the same screen coordinates,
because they run in the same per-monitor DPI context. This was verified at
150 % on one monitor.

### 5.9 Resolving the mail and its sender

1. Take `root = GetAncestor(wwgHwnd, GA_ROOT)`.
2. **Explorer:** compare `root` with `IOleWindow.GetWindow()` of
   `Application.ActiveExplorer()`. When they match:
   - with `ActiveInlineResponse` set, the click was in the inline editor. For
     the sender, use the original mail, `Selection(1)`;
   - otherwise use the reading pane mail, `Selection(1)`, when
     `Selection.Count == 1`.
3. **Inspector:** otherwise, walk `Application.Inspectors` with an index loop
   and compare each `IOleWindow` handle. Use the matching inspector's
   `CurrentItem`.
4. **Sender:** read `SenderName`. For the SMTP address:
   - `SenderEmailAddress` when `SenderEmailType == "SMTP"`;
   - otherwise `PropertyAccessor` `PR_SENDER_SMTP_ADDRESS`
     (`http://schemas.microsoft.com/mapi/proptag/0x5D01001F`).

   Put this in an adapter method (for example
   `IMailStore.GetSenderInfo(object item)`). If D2 adds one, reuse it.
5. **Rules:**
   - this reads envelope properties only, never the body;
   - release every COM object except the inline-response item, which belongs
     to Outlook;
   - if the item cannot be resolved, still warn, without the sender line. The
     banner came from the rendered text.

### 5.10 Warning dialog and "Open anyway"

**The dialog:**

- Content:
  - `PhishingLink_Title` as the title, then `PhishingLink_Text`;
  - the sender line;
  - the full URL in a selectable, wrapping text box;
  - "Shown in the email as" when `TextToDisplay` differs from the URL, a
    strong phishing signal.
- Buttons: **Open anyway** and **Cancel**. Cancel is the default and answers
  `Esc`.
- Owner: the `_WwG`'s root window, through `WindowInteropHelper.Owner`.
- It is shown through `Dispatcher.BeginInvoke`. Measured delay: 80–100 ms
  after the click.
- Set `_warningOpen` while it is shown. Clear `_warningPending` and
  `_warningOpen` in `finally`.

**Open anyway:**

1. Get the Word window again from the stored `_WwG` handle.
2. Re-acquire the link:
   - mouse path: `HitTest` at the stored point (R7);
   - Enter and menu paths: `SelectionLink` (*not tested*; same API).
3. If the URL is unchanged, call `Hyperlink.Follow()`. This keeps Outlook's own
   behaviour, including its security prompts for `file:` and unusual schemes.
4. Fallback, only if the link cannot be re-acquired (the mail closed, or the
   selection moved) **and** the scheme is http or https:
   `Process.Start(new ProcessStartInfo(url) { UseShellExecute = true })`.
   For any other scheme, log and do nothing.

**If the warning itself fails** (an exception while building or showing it),
log the error and follow the link: fail-open, as D2 does.

### 5.11 Installation, cleanup, threading and fail-open rules

- **Threading.** Everything runs on Outlook's UI thread: both hook
  procedures, the Word and Outlook object-model calls (in-process, same STA),
  and the dialog. There are no background threads and no waits.
- **Cleanup.** `Dispose` calls `UnhookWindowsHookEx` for both hooks. It is
  idempotent and runs from both `OnBeginShutdown` and `OnDisconnection`.
  - With fast shutdown, neither may run at exit (section 2.12). That is fine:
    thread hooks die with the thread.
  - At runtime, `OnDisconnection` must unhook before the add-in is released.
- **Fail-open everywhere:**
  - a hook error or COM failure passes the gesture;
  - a callback error leaves `cancelDefault = false`;
  - an unresolvable item still warns, without the sender line;
  - a failed warning follows the link.
- **Rules from the experiments:**
  - decide on the first sighting of a message;
  - expire the paired-up flag after 1.5 s and the pending-down key after 2 s;
  - never show UI or pump messages inside a hook;
  - keep both hook delegates in fields: Microsoft documents an
    `ExecutionEngineException` when the garbage collector moves a callback.
- **No signature, or setting off:** the hooks stay installed and pass
  everything through after the cheap checks. Installing or removing them when
  the setting changes would add state for no measurable gain.

### 5.12 Performance budget (from the measurements)

| Step | Budget | Measured |
|---|---|---|
| Mouse message that is not a click | one comparison | not instrumented |
| Click on a non-Word window | class-name lookup | negligible |
| Click on a `_WwG` | hit test | 0–4 ms |
| Click on a link | + rendered text + policy | 0–10 ms (about 45 ms on first use, spike, including one-off subscriptions) |
| Warning shown | after the hook returns | 80–100 ms after the click |

### 5.13 Logging and privacy

- Log one Information line per warning, with the surface (reading pane,
  inspector, inline, compose), `PhishingLinkFilter.HostForLog(url)` and the
  user's choice.
- **Never log the full URL:** paths and query strings can carry tokens
  (decision D-9).
- Log hook installation failures and the first error of each kind, then stay
  quiet.

### 5.14 Dev verification for D3 (e2e, by Claude)

- **Fake banner mail:** PostItem trick, convert to `IPM.Note`, then set
  `InternetCodepage = 65001` **on the converted MailItem** and rewrite
  `HTMLBody`. Without this, Outlook may store the banner in a character set
  that breaks the HTML match (section 2.11).
- **Link coordinates:** use the Word object model through `OBJID_NATIVEOM`
  and `Window.GetPoint`, from PowerShell (Appendix A).
- **Clicks:** real input, with the guarded click helper.
- **Check result:** watch the foreground window, and expect **no**
  `Example Domain` tab.
- **Cases:**
  - reading pane: click, Enter after Tab, context menu;
  - opened mail: the same three. Do this **after fixing D-6**, or the banner
    disappears on open;
  - reply in its own window: Ctrl+click;
  - inline reply: Ctrl+click;
  - image link and table "button" link;
  - a mail without banner (must open unchanged);
  - setting off;
  - no learned signature.
- **Cannot be checked on the dev machine:** real banner and Stormshield mail.
  Leave those checks open for the target, as section D says.

### 5.15 Documentation to update with D3

- `website/docs/user-guide/walkthrough.md`: the new setting row, and a short
  description of the warning.
- `docs/security.md` and `website/docs/security.md`:
  - two **thread-scoped Windows hooks on Outlook's own UI thread**: no global
    hook, no injection, no API patching;
  - the rendered mail text is read at click time;
  - no mail is modified;
  - the link host is logged;
  - known gap: touch.
- `website/docs/developer-guide/architecture.md`: the hook rules from section
  5.11.

---

## 6. Residual risks and target re-verification

1. **EDR and Stormshield reaction to `SetWindowsHookEx`.** The add-in makes
   two thread-scoped calls, with `hMod = NULL` and its own thread. Nothing was
   verified on the target.
   - **Check:** install the D3 build on the target, open banner mail, click
     links, and check the EDR console for alerts or blocks.
   - **Fallback:** if the hooks are flagged, subclass the `_WwG` windows
     (section 4).
2. **32-bit Outlook, Semi-Annual Enterprise Channel build.** Re-check:
   - the `_WwG` class name;
   - `OBJID_NATIVEOM` access;
   - `ProtectionType = 3` in received mail;
   - **acting on the button-down** (R5b);
   - the `HyperlinkOpen` repurpose with the toggle signature. An unknown
     `idMso` could make Office reject the whole custom UI: check that the
     RBLclass tab is present;
   - the `MSG`/`IntPtr` layout. The code uses `IntPtr.Size`.
3. **Touch-screen laptops.** Taps open links without a warning (P1, P2),
   unless D-2 adds the warning after the tap.
4. **Stormshield-encrypted banner mail.**
   - Is the banner inside the decrypted content that Word renders? A
     transport rule cannot change an encrypted body; it may wrap the message
     or skip the banner.
   - Check that reading the rendered text and the sender properties causes no
     smart-card prompt.
5. **Real banner mail.** Run the existing Settings diagnostic on banner mails
   from several senders and character sets (section 2.11). It shows how often
   the HTML match fails, which bears on D-3 and on D2.
6. **Not tested anywhere:**
   - the context menu opened from the keyboard (Shift+F10);
   - the context menu in a reply window;
   - "Open anyway" through `SelectionLink().Follow()`;
   - multi-monitor setups with mixed DPI;
   - very long links wrapping across lines (only the boundary branch depends
     on the box);
   - links inside tables nested in quoted text.
7. **Future Office changes.** If Word stops rendering mail in `_WwG` windows,
   the guard sees no Word clicks and fails open: no warnings, no breakage.
   - Log once per session, at the first `_WwG` click, so a silent regression
     shows in the logs.
8. **Interaction with the existing banner strip.** Until D-6 is fixed,
   opening a banner mail with *Strip banner from replies* on removes its
   banner, and no warning can fire for it afterwards.

---

## 7. Decisions needed from the user

| # | Decision | Options | Recommendation |
|---|---|---|---|
| D-1 | Mechanism for D3 | (a) **before the open**: `WH_MOUSE` + `WH_GETMESSAGE` (Enter) + repurposed `HyperlinkOpen`; (b) after the open only (foreground watching); (c) drop D3 as not feasible | **(a)**, subject to the EDR check of section 6.1 |
| D-2 | Touch taps (cannot be intercepted) | (a) accept and document the gap; (b) warn **after** the tap with the URL: the `WH_GETMESSAGE` hook sees where the finger went down, hit-tests that point, and posts `PhishingLink_AfterTouch_Text` when the contact ends as a tap | **(a) for D3**, (b) as a later item if touch devices are common on the target |
| D-3 | Banner test used at click time | (a) **rendered text** (letters and digits): survives character-set differences, needs no item access, no smart-card prompt; (b) HTML `ContainsBanner` (D2's definition); (c) either one | **(a) for D3**, and consider (c) for D2, which reads the HTML anyway |
| D-4 | `mailto:` links | (a) no warning (the send-time guard already covers writing to outside); (b) warn | **(a)** |
| D-5 | Context menu "Open Hyperlink" | (a) repurpose `HyperlinkOpen` (verified, cancellable, menu kept); (b) suppress the menu on banner links (`WindowBeforeRightClick`, removes "Copy Hyperlink"); (c) leave it uncovered | **(a)**, with (b) as fallback if (a) fails on the target |
| D-6 | Existing defect: banner stripped from received mail opened in its own window | (a) fix it now, as a separate item before D2/D3 (strip only when `mail.Sent == false`); (b) leave it | **(a)**: it modifies received mail and blinds D2/D3 |
| D-7 | Ctrl+click in replies (own window and inline) | (a) include (same code path); (b) received mail only | **(a)**. Inline replies are checked against the original mail. Reply windows are covered only while the draft still contains the banner, which depends on *Strip banner from replies* |
| D-8 | EDR de-risking order | (a) first D3 build to the target for an EDR check before pilot; (b) ship and see | **(a)** |
| D-9 | What is logged for a warning | (a) scheme and host only; (b) full URL; (c) nothing | **(a)** |

---

## Appendix A: how the experiments were run

All prototype code was throwaway.

- **Spike class:** one file in the add-in, created in `OnStartupComplete`,
  plus temporary edits to the csproj (Word interop and UI Automation
  references, for convenience only) and to `Ribbon.fr.xml`.
- **Clean-up:** the spike files and edits were reverted and the clean build
  reloaded. The fake mails and their folder were deleted, and the temporary
  reading-pane preference was restored.
- **No helper scripts were kept** in the repository. The techniques below are
  enough to rebuild them for D3's verification.

The techniques:

- **Safe real clicks.**
  1. Make the script process per-monitor DPI aware
     (`SetProcessDpiAwarenessContext(-4)`), so every coordinate is physical.
  2. Bring the target top-level window to the front: set it topmost, call
     `SetForegroundWindow`, then remove topmost.
  3. **Click only if `GetAncestor(WindowFromPoint(pt), GA_ROOT)` is the
     target.**
  4. Send `SetCursorPos`, then `mouse_event` down and up. For Ctrl+click, wrap
     them in `keybd_event(VK_CONTROL)`.
- **Link positions.**
  1. Find the `_WwG` with `EnumChildWindows`.
  2. Call `AccessibleObjectFromWindow(hwnd, 0xFFFFFFF0, IID_IDispatch)` from
     PowerShell; it works across processes.
  3. Walk `Document.Hyperlinks`, calling `Window.GetPoint(ref l, ref t, ref w,
     ref h, hyperlink.Range)` on each one.
- **Touch.** `InitializeTouchInjection(1, 1)`, then `InjectTouchInput` three
  times: down with in-range and in-contact flags, an update, then up. No admin
  rights needed.
- **Observing.** A second PowerShell process logged:
  - `SetWinEventHook` out of context: `EVENT_SYSTEM_FOREGROUND` for every
    process, and every event from `OUTLOOK.EXE`;
  - UI Automation event handlers;
  - WMI process creation.

  Success meant **no** foreground change to the browser and no new
  `Example Domain` tab.
- **Fake mails.** `Items.Add(6)` (a PostItem), `HTMLBody`, `Save()`, then
  `MessageClass = "IPM.Note"` and `Save()`. Reopen the item, set
  `InternetCodepage = 65001`, rewrite `HTMLBody`, and set `UnRead`. The image
  link used an attachment with `PR_ATTACH_CONTENT_ID` and a `cid:` source.
- **Spike configuration.** An ini file re-read at each click, so one build
  served all experiments. Settings: `mode`, `hittest`, `banner`,
  `requireBanner`, `after`, `follow`, `gmMouse`, `gmEnter`, `gmPointer`,
  `gmPointerPeek`, `cancelMenuOpen`, `cancelRightClick`.

---

## Sources

Microsoft documentation:

- [SetWindowsHookExW function](https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-setwindowshookexw):
  thread versus global scope, `hMod = NULL` for a thread of the current
  process, and keeping .NET callbacks alive.
- [MouseProc callback (WH_MOUSE)](https://learn.microsoft.com/en-us/windows/win32/winmsg/mouseproc).
- [GetMsgProc callback (WH_GETMESSAGE)](https://learn.microsoft.com/en-us/windows/win32/winmsg/getmsgproc):
  "The GetMsgProc hook procedure can examine or modify the message", and
  `PM_NOREMOVE` / `PM_REMOVE`.
- [SetWinEventHook function](https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-setwineventhook).
- [AccessibleObjectFromWindow function](https://learn.microsoft.com/en-us/windows/win32/api/oleacc/nf-oleacc-accessibleobjectfromwindow),
  including `OBJID_NATIVEOM`.
- [UI Automation threading issues](https://learn.microsoft.com/en-us/windows/win32/winauto/uiauto-threading).
- [InjectTouchInput function](https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-injecttouchinput).
- Word object model:
  - [Window.RangeFromPoint](https://learn.microsoft.com/en-us/office/vba/api/word.window.rangefrompoint)
  - [Window.GetPoint](https://learn.microsoft.com/en-us/office/vba/api/word.window.getpoint)
  - [Hyperlink.Follow](https://learn.microsoft.com/en-us/office/vba/api/word.hyperlink.follow)
  - [WdProtectionType](https://learn.microsoft.com/en-us/office/vba/api/word.wdprotectiontype)
  - [Options.CtrlClickHyperlinkToOpen](https://learn.microsoft.com/en-us/office/vba/api/word.options.ctrlclickhyperlinktoopen)
  - [Application.WindowSelectionChange](https://learn.microsoft.com/en-us/office/vba/api/word.application.windowselectionchange)
  - [Application.WindowBeforeRightClick](https://learn.microsoft.com/en-us/office/vba/api/word.application.windowbeforerightclick)
- Outlook object model:
  - [Inspector.WordEditor](https://learn.microsoft.com/en-us/office/vba/api/outlook.inspector.wordeditor)
  - [Explorer.ActiveInlineResponseWordEditor](https://learn.microsoft.com/en-us/office/vba/api/outlook.explorer.activeinlineresponsewordeditor)
- [Temporarily Repurpose Commands on the Office Fluent Ribbon](https://learn.microsoft.com/en-us/previous-versions/office/developer/office-2007/bb462633(v=office.12)):
  `<commands>/<command idMso>` with an `onAction(control, ref cancelDefault)`
  callback. The toggle form `(control, pressed, ref cancelDefault)` that
  `HyperlinkOpen` needs was established by experiment (X1–X3).

Project:

- `ROADMAP.md` section D (D1–D3)
- `src/RBLclass.Core/ExternalBannerStripper.cs` (`ContainsBanner`, `Strip`)
- `src/RBLclass.AddIn/RblClassAddIn.cs` (`OnNewInspector` at line 754,
  `OnBeginShutdown` at 642, `OnDisconnection` at 121)
- `src/RBLclass.Outlook.Adapter/OutlookMailStore.cs`
  (`StripBannerFromDraft` at line 1816, `StripBannerOnLiveMail` at 1827)
- `docs/security.md`
