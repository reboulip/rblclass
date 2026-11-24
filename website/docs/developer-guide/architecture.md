# Architecture

RBLclass is **strictly layered**. Dependencies point downward only.

```text
RBLclass.AddIn              (.NET FW 4.8)  COM add-in shell:
      │                                    IDTExtensibility2, IRibbonExtensibility,
      │                                    ribbon callbacks, Custom Task Panes,
      │                                    Outlook event subscriptions
RBLclass.Outlook.Adapter    (.NET FW 4.8)  COM access to the Outlook Object Model,
      │                                    implements RBLclass.Core interfaces
      ▼
RBLclass.Core               (.NET Standard 2.0)  business logic — no Outlook,
                                             no UI, no I/O except via injected
                                             interfaces
```

!!! abstract "The one rule that matters most"
    `RBLclass.Core` must **never** reference
    `Microsoft.Office.Interop.Outlook`, `System.Windows`, or the COM-add-in shell
    assemblies. That is what keeps the business logic portable on the day the COM
    add-in model has to be left behind.

When a feature needs Outlook data, the flow is always the same:

1. Define an **interface** in `RBLclass.Core`.
2. Implement it in `RBLclass.Outlook.Adapter` (the COM boundary).
3. Write the business logic in `RBLclass.Core` against the interface.
4. Wire it up from `RBLclass.AddIn`.

## Folder layout

```text
/src
  /RBLclass.Core              business logic, .NET Standard 2.0
  /RBLclass.Outlook.Adapter   COM adapter, .NET Framework 4.8
  /RBLclass.AddIn             COM add-in shell, .NET Framework 4.8
/installer                    WiX MSI project (Phase 1)
/tests
  /RBLclass.Core.Tests        xUnit
/website                      this documentation site (MkDocs Material)
/docs                         architecture notes, install guide, roadmap detail
ROADMAP.md  CLAUDE.md  README.md
```

Files most often relevant when changing behaviour:

- `src/RBLclass.AddIn/RblClassAddIn.cs` — COM entry points, ribbon callbacks, CTP
  management, event registration, startup.
- `src/RBLclass.AddIn/ViewModels/MainPaneViewModel.cs` — main pane logic.
- `src/RBLclass.AddIn/Views/MainPaneView.xaml` (+ `.xaml.cs`) — pane XAML.
- `src/RBLclass.Core/Settings.cs` + `SettingsKeys.cs` — typed settings.
- `src/RBLclass.AddIn/Ribbon.xml` (+ `.fr.xml`, `.de.xml`) — ribbon definition.

## COM interop rules (shell only)

These conventions exist because getting them wrong takes Outlook down silently
and is very hard to diagnose:

- **Never hand-roll `[ComImport]` declarations** of `IDTExtensibility2`,
  `IRibbonExtensibility`, `IRibbonControl`, etc. Reference the canonical PIAs
  from the GAC (`Extensibility`, `office`,
  `Microsoft.Office.Interop.Outlook`) with `<Private>false</Private>`. The CLR
  auto-generates marshalling stubs from interface metadata; subtle deviations
  from the canonical PIAs produce stubs that access-violate inside
  `IL_STUB_COMtoCLR`, surfacing as `ExecutionEngineException` (HRESULT
  `0x80131506`).
- **`[ClassInterface(ClassInterfaceType.AutoDispatch)]` on the add-in class is
  mandatory** — not `None`. Office resolves ribbon `onAction` callbacks via
  `IDispatch::GetIDsOfNames` on the class itself.
- **Do not embed Office interop types** — use direct GAC references, which are
  never copied to the output.
- **One narrow exception: Extended MAPI in the adapter.** Extended MAPI has no
  PIA, so `MapiMessageMover` calls exactly two `IMAPIFolder` methods,
  `CopyMessages` (vtable slot 20) and `DeleteMessages` (slot 21), through their
  vtable slots with explicit `StdCall` delegates. The pointer comes from
  `Folder.MAPIOBJECT` plus `QueryInterface` for `IID_IMAPIFolder`
  (`{0002030C-0000-0000-C000-000000000046}`). There is still **no `[ComImport]`
  interface**, so the rule above stands. `MAPIOBJECT` is in-process and
  UI-thread only. `ENTRYLIST`/`SBinary` are declared with `IntPtr` fields, so the
  same AnyCPU binary works in 32-bit and 64-bit Outlook. Do not add further
  hand-rolled MAPI calls without the same justification.

## COM lifetime (adapter only)

Outlook COM objects must be released explicitly, or Outlook leaks and crashes:

- Wrap every COM object in a `ComRef<T> : IDisposable` that calls
  `Marshal.ReleaseComObject` in `Dispose`.
- **Never chain property accesses** on COM objects — `mail.Parent.Parent.Name`
  leaks the intermediate `Parent` objects.
- Loop over `Items` / `Folders` / `Stores` with a `for` index (not `foreach`),
  disposing each element.
- `MapiMessageMover` calls `Marshal.Release` on every pointer it obtains and
  frees every `HGlobal` in a `finally` block; it never throws and logs failures
  at Warning with the HRESULT.

### Protected mail is filed at store level

`MailItem.Move`/`Copy`/`Delete` persist an item's in-memory state, and a
mail-encryption product (Stormshield) decrypts the opened item in memory, so
filing a protected mail through the object model stores it in clear. In
`OutlookMailStore`, when the resolved item is a `MailItem` that reads as
protected (or was filed at store level earlier in the session), the item
reference is released and `Move`, `Copy`, `DeleteItem` and
`DeleteItemPermanently` go through `MapiMessageMover` instead. There is no
object-model fallback: a failed transfer throws and the mail stays where it was.
The new EntryID is found by diffing the destination's contents table before and
after (protected-class filter, narrowed on Internet Message-ID, else subject).

A PST move keeps the EntryID, and Outlook can hand back its stale in-memory
item. Refs from a store-level filing are therefore `MailItemRef.VerifiedProtected`
(`ClassifierService` skips its object-model post-filing check for them), and a
session map of store-level filings keeps them recognised as protected. If the new
row is not identified, the ref is `MailItemRef.Untracked` (`IsTracked` is false):
counted as filed, never added to Undo, never used for a safety copy.

## Threading

- The Outlook OM is single-threaded apartment (STA); **all COM calls happen on
  the main thread**.
- Background work (indexing, SQLite I/O, HTTP) runs on the thread pool.
- A `SynchronizationContext` captured at add-in startup marshals results back to
  the Outlook UI thread.
- Never block the UI thread on COM calls inside loops — batch and yield.
- The PDF export of *Save as file* follows this rule: Outlook saves each mail to
  a temporary MHTML file on the UI thread, then Word (late-bound, out of process)
  converts the files on a **dedicated STA thread**. Outlook OM calls stay on the
  UI thread.

## Storage & SQLite policy

- One connection per logical operation, opened with **WAL** journal mode.
- All schema changes go through **versioned migrations** (`SchemaVersion`
  table) — never alter the schema in-place.
- The **FTS5** virtual table for search is kept separate from the main `Mails`
  table so it can be rebuilt without data loss.
- EntryIDs can change after some Outlook operations (e.g. moving between stores).
  Cache by `(StoreID, EntryID)` and **tolerate misses**.

## Outlook events

- Register handlers in `OnStartupComplete`, unregister in `OnBeginShutdown`.
- Keep references to event source objects in **fields** — otherwise GC collects
  them and events stop firing silently.
- Item-level events (`BeforeAttachmentWriteToTempFile`, `BeforeAttachmentPreview`,
  used by the external-attachment phishing warning) follow the same rule: keep the
  hooked item and its delegates in fields, hook the single selected mail (re-hooked
  from the debounced selection tick) and each opened mail window's item (hooked on
  its first `Inspector.Activate`), and release the hooks on change, on
  `Inspector.Close` and at shutdown.
- The external-link phishing warning uses two **thread-scoped** Windows hooks on
  Outlook's own UI thread (`WH_MOUSE` for clicks, `WH_GETMESSAGE` for ++enter++),
  installed with `hMod` NULL: never global, nothing injected, no API patched. It
  also repurposes the built-in ribbon command `HyperlinkOpen` through a
  `<commands>` block in the three Ribbon XML files; the public callback must keep
  exactly the signature `OnHyperlinkOpenRepurposed(IRibbonControl, bool pressed,
  ref bool cancelDefault)`. The Word object model is reached late-bound through
  `AccessibleObjectFromWindow` (`OBJID_NATIVEOM`); the only hand-declared COM
  interface is the standard `IOleWindow`. Hook rules:
    - hold the hook delegates in **fields**: a collected callback crashes the
      process;
    - never show UI or pump messages inside a hook: post the warning so it runs
      after the hook returns;
    - always call `CallNextHookEx`;
    - unhook in both `OnBeginShutdown` and `OnDisconnection`. A fast shutdown may
      skip `OnBeginShutdown`, which is harmless because thread hooks die with the
      process.
- Handle `StoreRemove` gracefully: PST stores can be closed by the user at any
  time.

## Error handling

- Outlook COM calls throw `COMException`. **Catch and log; never let an exception
  escape into Outlook** — it disables the add-in by flipping `LoadBehavior` from
  3 to 2 in the registry.
- Top-level handlers wrap every ribbon callback, every event handler, and every
  fire-and-forget `Task`.

## Localization

- XAML strings go through `{loc:Loc Key=...}`
  (`RBLclass.AddIn.Localization.LocExtension`).
- Code-behind / ViewModel strings go through
  `ILocalizationService.GetString` / `GetString(key, args)`.
- Count-dependent strings use a `_One` / `_Other` key pair resolved via
  `ILocalizationService.Plural(...)` — never concatenate, because French and
  German plural forms differ from English.
- Ribbon buttons/tooltips need matching entries in all three `Ribbon*.xml`
  files. `ResourceParityTests` enforces parity across `.resx` and ribbon files.

## Re-encrypted filing (experimental)

An opt-in path (setting `ReencryptProtectedOnFile`, off by default) that files a
protected mail that is readable decrypted in memory as a freshly encrypted copy.
It follows the usual layering.

**`RBLclass.Core`** (no Outlook, no crypto API):

- `IMailRebuilder` — the adapter contract: `ReadDecryptedContent`, `Encrypt`,
  `CreateEncryptedCopy`, `VerifyEncryptedCopy`, `ResolveCertificate`,
  `ListEncryptionCertificates`. It never throws. The COM members run on the STA
  thread; `Encrypt` and `ListEncryptionCertificates` are COM-free.
- `MailContent` / `MailPart` — the rebuilt-content model (body, body images,
  attachments, envelope).
- `MimeEntityBuilder` — builds the RFC 2045/2046 message: multipart
  alternative/related/mixed, base64, RFC 2231/2047 file names, a meta-charset
  rewrite to utf-8, streaming output and boundary-collision checks.
- `ReencryptContentPlanner` — pure logic that applies the attachment,
  body-image, label and banner rules to the rebuilt content, mirroring the
  in-place `ApplyAttachmentRemoval` rules.
- `ReencryptEligibility` and `ReencryptOutcome` — which mail takes the path, and
  the per-item result (re-encrypted, not applicable, or fallback with a reason).
- `EncryptionCertificatePolicy`, `CertificateCandidate`, `CertificateInfo` — the
  certificate filtering and *Automatic* selection rules.
- `ClassifierService` takes an optional `IMailRebuilder`; with
  `ClassifyOptions.ReencryptProtected`, its protected-item branch tries the
  re-encrypted path before the standard loop, and falls back to it on any failure.
  The path is gated on `ChangesContent(request, item)`: `RemoveAttachments`, or
  `StripBanner` (a banner signature is set), or any `AttachmentDisposition` row
  for this item whose action is not Keep. A plain classify never touches
  `IMailRebuilder` and goes through store-level filing.

**`RBLclass.Outlook.Adapter`**:

- `OutlookMailRebuilder` creates the `IPM.Note.SMIME` item with a "PostItem
  trick": `Items.Add(olPostItem)`, add a `smime.p7m` attachment tagged
  `PR_ATTACH_MIME_TAG` `application/pkcs7-mime`, flip `MessageClass` to
  `IPM.Note` and reopen as a `MailItem`, copy the recipients and envelope MAPI
  properties, set `PR_ICON_INDEX` to -1, then flip to `IPM.Note.SMIME`. The
  payload is byte-checked before the flip and the stub is deleted on failure.
- `MailItemProbes` — probes shared with `OutlookMailStore` (moved out of it).
- `CmsEnvelopeBuilder` (Core) builds the CMS EnvelopedData DER itself: pure
  managed, `netstandard2.0`, no new dependency (AES-256-CBC, RSA PKCS#1 v1.5 key
  transport, one `KeyTransRecipientInfo` by issuer + serial, like a real
  Stormshield mail). `EnvelopedCms` on .NET Framework 4.8 can only emit RSA-OAEP.
  The adapter passes the certificate's issuer name DER, serial (big-endian) and
  RSA public key, then keeps a decode-only `EnvelopedCms` self-check that logs
  cipher and key transport at Debug and warns if the key transport is not
  `rsaEncryption`.
- **Step breadcrumbs.** Before each step `ClassifierService.TryReencrypt` logs
  `Re-encrypted filing step {Step} ({EntryId})` (eligibility, certificate, read,
  build, encrypt, then create and verify once per destination, and original only
  when the original goes to Deleted Items). `OutlookMailRebuilder.CreateEncryptedCopy`
  logs `Re-encryption step {Step} ({EntryId}).` (stub, reopen, payload-check,
  envelope, class-switch, and `saved` after the final save). They are
  **Information**, not Debug, because target workstations run without Diagnostic
  logging; the Serilog file sink is unbuffered, so after an Outlook crash the last
  line names the step that hung.

See [Security](../security.md) for the trust and plaintext-handling implications.

## Performance targets

| Operation | Target |
| --- | --- |
| Folder search | < 100 ms for a query over 100k indexed mails |
| Classify action | Optimistic UI, perceived latency < 300 ms (the actual Move may complete asynchronously) |
| Add-in startup | Contributes < 500 ms to Outlook startup; initial PST indexing runs in the background afterwards |

## Outlook OM quirks worth remembering

- `MailItem.Move` returns the moved item **in the destination store**; the
  original COM reference becomes invalid.
- `Explorer.Selection` can contain non-`MailItem` objects (`MeetingItem`,
  `ReportItem`, …) — always check `is MailItem` before casting.
- `Folder.Items.Find` / `Restrict` is much faster than iterating, but uses its
  own DASL/Jet query syntax — document any non-trivial query.

The full, authoritative version of these rules lives in
[`CLAUDE.md`](https://github.com/reboulip/rblclass/blob/main/CLAUDE.md) at the
repo root.
