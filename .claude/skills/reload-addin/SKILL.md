---
name: reload-addin
description: Local dev loop for this machine - rebuild the RBLclass product add-in and reload it into Outlook (build, close Outlook, install, restart Outlook) so a freshly-developed change can be verified live. Run this at the end of a development task when the user needs to check the change in Outlook. Targets the real product (src/RBLclass.AddIn), NOT the POC.
---

# /reload-addin — rebuild and reload the add-in into Outlook

The fast inner dev loop on this (the only test) machine. It rebuilds
`src/RBLclass.AddIn`, swaps the freshly-built binary into Outlook's
per-user install location, and restarts Outlook so the change can be
verified live. It does **not** package an install kit — that is
`/make-release`.

## When to run it

Run this **at the end of a development task**, once the work builds and you
want the user to confirm the behaviour in Outlook matches expectations.
The user has standing authorisation for Claude to trigger this loop in that
situation. Be aware it **closes the user's Outlook** — give them a heads-up
in the same turn, and don't run it mid-task or while they may have unsaved
Outlook work without flagging it.

## Run

From the repo root:

```powershell
.\.claude\skills\reload-addin\Reload-Addin.ps1            # Debug, restarts Outlook
.\.claude\skills\reload-addin\Reload-Addin.ps1 -Clean     # Rebuild (clean)
.\.claude\skills\reload-addin\Reload-Addin.ps1 -NoStart   # don't relaunch (e.g. to attach a debugger)
.\.claude\skills\reload-addin\Reload-Addin.ps1 -Configuration Release
```

Order is deliberate:

1. **Build** `src/RBLclass.AddIn` (`Debug|AnyCPU` by default) — Outlook
   stays open, so a **broken build never closes your Outlook**.
2. **Close** Outlook gracefully; force-kill only after `-CloseTimeoutSeconds`
   (default 20s) if a dialog blocks the graceful exit.
3. **Install** the build output via the make-release installer
   ([Install-RblClassAddIn.ps1](../make-release/Install-RblClassAddIn.ps1)
   `-SkipSigning`), copying to `%LocalAppData%\RBLclass\AddIn` and writing
   the HKCU COM + Outlook registry entries.
4. **Restart** Outlook (unless `-NoStart`).

## Shared identity

This skill reuses the single source of truth from `/make-release`:
[release.config.json](../make-release/release.config.json) and its
installer. There is **no second COM identity** — the dev loop and the
shippable kit register the add-in under the exact same CLSID/ProgId, so
"works after reload" means "works after install from a kit". The same
all-zeros sentinel gate applies: the loop refuses to run until the config
is reconciled with the real `RBLclass.AddIn` `[Guid]`/`[ProgId]`.

## Preconditions

- Targets the **real product** only. If `RBLclass.sln` does not exist yet
  (Phase 1 not built), the loop fails clearly and changes nothing — it
  does not fall back to the POC. To iterate on the POC, use its own
  [poc/scripts/Install-HelloPstAddIn.ps1](../../../poc/scripts/Install-HelloPstAddIn.ps1).
- Signing is skipped (`-SkipSigning`) — not required to load (Phase 0).
- Visual Studio 2022 (MSBuild via vswhere) + the standalone .NET SDK on
  PATH.

## After it runs

Tell the user the add-in was reloaded and Outlook is restarting, and what
to look for (the RBLclass ribbon tab and the specific behaviour you just
changed), so they can confirm it matches expectations.
