using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Windows.Threading;
using RBLclass.AddIn.ViewModels;
using RBLclass.AddIn.Views;
using RBLclass.Core;
using RBLclass.Core.Persistence;
using Serilog;

namespace RBLclass.AddIn.Phishing
{
    /// <summary>
    /// Warns BEFORE a link from a mail carrying the organization's external-sender banner
    /// opens (v2.13.0.0 D3; mechanism validated in docs/phishing-link-study.md section 5).
    ///
    /// Two thread-scoped hooks on Outlook's own UI thread (never global, hMod NULL, nothing
    /// injected): WH_MOUSE for clicks and WH_GETMESSAGE for Enter on a link; plus the
    /// repurposed ribbon command HyperlinkOpen for the context menu. A click is swallowed
    /// when it would follow a link in a banner mail; the warning is shown after the hook
    /// returns (never UI inside a hook). "Open anyway" re-follows the link through
    /// Word's own Hyperlink.Follow(). Fail-open everywhere. The banner is detected from
    /// the rendered text at click time; no mail is modified; only scheme://host is logged.
    /// Touch taps never reach WH_MOUSE and cannot be blocked (accepted gap).
    /// </summary>
    internal sealed class LinkClickGuard : IDisposable
    {
        private const int SwallowUpExpiryMs = 1500;
        private const int PendingKeyExpiryMs = 2000;
        private const int RightClickMemoryMs = 30000;
        private const string WordDocumentWindowClass = "_WwG";

        private enum Source { Mouse, Enter, Menu }

        private sealed class PendingWarning
        {
            public IntPtr Hwnd;
            public Source Source;
            public int X;
            public int Y;
            public string Url;
            public string DisplayedText;
        }

        private readonly ISettingsStore _settings;
        private readonly LinkMailResolver _resolver;
        private readonly Dispatcher _dispatcher;

        // The delegates are held in fields for the add-in's lifetime: a collected hook
        // callback crashes the process (ExecutionEngineException).
        private readonly NativeMethods.HookProc _mouseProc;
        private readonly NativeMethods.HookProc _getMsgProc;
        private IntPtr _mouseHook = IntPtr.Zero;
        private IntPtr _getMsgHook = IntPtr.Zero;
        private bool _disposed;

        private volatile bool _warningPending;
        private volatile bool _warningOpen;

        private bool _swallowUp;
        private int _swallowUpSince;

        private bool _hasKey;
        private IntPtr _keyHwnd;
        private int _keyX, _keyY;
        private int _keySince;
        private bool _keyDecision;

        private IntPtr _lastRightWwg = IntPtr.Zero;
        private int _lastRightTick;

        private bool _loggedHookError;

        public LinkClickGuard(ISettingsStore settings, LinkMailResolver resolver)
        {
            _settings = settings;
            _resolver = resolver;
            _dispatcher = Dispatcher.CurrentDispatcher;
            _mouseProc = MouseProc;
            _getMsgProc = GetMsgProc;
        }

        /// <summary>Install both thread-scoped hooks. Failures are logged and leave the guard inert.</summary>
        public void Install()
        {
            uint threadId = NativeMethods.GetCurrentThreadId();

            _mouseHook = NativeMethods.SetWindowsHookEx(NativeMethods.WH_MOUSE, _mouseProc, IntPtr.Zero, threadId);
            if (_mouseHook == IntPtr.Zero)
                Log.Warning("Link warning: WH_MOUSE hook not installed (Win32 error {Error}).", Marshal.GetLastWin32Error());

            _getMsgHook = NativeMethods.SetWindowsHookEx(NativeMethods.WH_GETMESSAGE, _getMsgProc, IntPtr.Zero, threadId);
            if (_getMsgHook == IntPtr.Zero)
                Log.Warning("Link warning: WH_GETMESSAGE hook not installed (Win32 error {Error}).", Marshal.GetLastWin32Error());

            Log.Information("Link warning guard installed (mouse={Mouse}, enter={Enter}).",
                _mouseHook != IntPtr.Zero, _getMsgHook != IntPtr.Zero);
        }

        public void Dispose()
        {
            if (_disposed) return;
            _disposed = true;
            if (_mouseHook != IntPtr.Zero)
            {
                try { NativeMethods.UnhookWindowsHookEx(_mouseHook); } catch { }
                _mouseHook = IntPtr.Zero;
            }
            if (_getMsgHook != IntPtr.Zero)
            {
                try { NativeMethods.UnhookWindowsHookEx(_getMsgHook); } catch { }
                _getMsgHook = IntPtr.Zero;
            }
        }

        // ---- Mouse path ----------------------------------------------------------

        private IntPtr MouseProc(int nCode, IntPtr wParam, IntPtr lParam)
        {
            try
            {
                if (nCode < 0 || _disposed) return Next(_mouseHook, nCode, wParam, lParam);

                int msg = (int)wParam;
                if (msg != NativeMethods.WM_LBUTTONDOWN && msg != NativeMethods.WM_LBUTTONDBLCLK
                    && msg != NativeMethods.WM_LBUTTONUP && msg != NativeMethods.WM_RBUTTONDOWN)
                    return Next(_mouseHook, nCode, wParam, lParam); // hot path: mouse moves and the rest

                if (msg == NativeMethods.WM_LBUTTONUP)
                {
                    if (_swallowUp)
                    {
                        _swallowUp = false;
                        // The matching up can be dropped by Windows when a modal opens: expire it.
                        if (Elapsed(_swallowUpSince) < SwallowUpExpiryMs) return (IntPtr)1;
                    }
                    return Next(_mouseHook, nCode, wParam, lParam);
                }

                var mhs = Marshal.PtrToStructure<NativeMethods.MOUSEHOOKSTRUCT>(lParam);

                if (msg == NativeMethods.WM_RBUTTONDOWN)
                {
                    if (IsWordDocumentWindow(mhs.hwnd))
                    {
                        _lastRightWwg = mhs.hwnd;
                        _lastRightTick = Environment.TickCount;
                    }
                    return Next(_mouseHook, nCode, wParam, lParam);
                }

                // One decision per physical click: the same message can be seen twice
                // (peek at HC_NOREMOVE, then HC_ACTION). Returning 1 at a peek removes it.
                bool swallow;
                bool sameKey = _hasKey && _keyHwnd == mhs.hwnd && _keyX == mhs.pt.X && _keyY == mhs.pt.Y
                               && Elapsed(_keySince) < PendingKeyExpiryMs;
                if (sameKey)
                {
                    swallow = _keyDecision;
                }
                else
                {
                    swallow = DecideMouse(mhs.hwnd, mhs.pt.X, mhs.pt.Y);
                    _hasKey = true;
                    _keyHwnd = mhs.hwnd;
                    _keyX = mhs.pt.X;
                    _keyY = mhs.pt.Y;
                    _keySince = Environment.TickCount;
                    _keyDecision = swallow;
                }
                if (nCode == 0) _hasKey = false; // HC_ACTION: this click is finished

                return swallow ? (IntPtr)1 : Next(_mouseHook, nCode, wParam, lParam);
            }
            catch (Exception ex)
            {
                LogHookErrorOnce(ex);
                return Next(_mouseHook, nCode, wParam, lParam);
            }
        }

        private bool DecideMouse(IntPtr hwnd, int x, int y)
        {
            if (!IsWordDocumentWindow(hwnd)) return false;

            // A warning is up or queued: the second click of a double-click, or an impatient one.
            if (_warningPending || _warningOpen) return true;

            var settings = Settings.Load(_settings);
            string signature = settings.ExternalBannerSignature;
            if (!settings.PhishingWarnLinks || string.IsNullOrWhiteSpace(signature)) return false;

            using (var win = WordWindow.TryGet(hwnd))
            {
                if (win == null) return false;

                // Would this click follow the link? A plain click in a reply only moves the caret.
                bool follows = win.IsReadOnly()
                    || !win.CtrlClickRequired()
                    || NativeMethods.GetKeyState(NativeMethods.VK_CONTROL) < 0;
                if (!follows) return false;

                var link = win.HitTest(x, y);
                if (link == null || !PhishingLinkFilter.IsWarnable(link.Url)) return false;

                if (!PhishingWarningPolicy.ShouldWarnLink(true, signature, win.RenderedText(), link.Url))
                    return false;

                _swallowUp = true;
                _swallowUpSince = Environment.TickCount;
                PostWarning(new PendingWarning
                {
                    Hwnd = hwnd, Source = Source.Mouse, X = x, Y = y,
                    Url = link.Url, DisplayedText = link.TextToDisplay
                });
                return true;
            }
        }

        // ---- Enter path ----------------------------------------------------------

        private IntPtr GetMsgProc(int nCode, IntPtr wParam, IntPtr lParam)
        {
            try
            {
                if (nCode < 0 || _disposed || (int)wParam != NativeMethods.PM_REMOVE)
                    return Next(_getMsgHook, nCode, wParam, lParam);

                int message = Marshal.ReadInt32(lParam, IntPtr.Size);
                if (message != NativeMethods.WM_KEYDOWN)
                    return Next(_getMsgHook, nCode, wParam, lParam); // hot path

                var msg = Marshal.PtrToStructure<NativeMethods.MSG>(lParam);
                if ((int)msg.wParam != NativeMethods.VK_RETURN || !IsWordDocumentWindow(msg.hwnd))
                    return Next(_getMsgHook, nCode, wParam, lParam);

                if (NativeMethods.GetKeyState(NativeMethods.VK_CONTROL) < 0
                    || NativeMethods.GetKeyState(NativeMethods.VK_MENU) < 0
                    || NativeMethods.GetKeyState(NativeMethods.VK_SHIFT) < 0)
                    return Next(_getMsgHook, nCode, wParam, lParam);

                // Auto-repeat of the Enter that started the warning.
                if (_warningPending || _warningOpen)
                {
                    Neutralize(lParam);
                    return Next(_getMsgHook, nCode, wParam, lParam);
                }

                if (DecideEnter(msg.hwnd)) Neutralize(lParam);
            }
            catch (Exception ex)
            {
                LogHookErrorOnce(ex);
            }
            return Next(_getMsgHook, nCode, wParam, lParam);
        }

        private bool DecideEnter(IntPtr hwnd)
        {
            Log.Debug("Link warning: Enter pressed in a Word document window.");
            var settings = Settings.Load(_settings);
            string signature = settings.ExternalBannerSignature;
            if (!settings.PhishingWarnLinks || string.IsNullOrWhiteSpace(signature)) return false;

            using (var win = WordWindow.TryGet(hwnd))
            {
                // In a reply Enter inserts a paragraph: only a read-only document follows links on Enter.
                if (win == null || !win.IsReadOnly()) return false;

                var link = win.SelectionLink();
                Log.Debug("Link warning: Enter, selection link found: {Found}.", link != null);
                if (link == null || !PhishingLinkFilter.IsWarnable(link.Url)) return false;
                if (!PhishingWarningPolicy.ShouldWarnLink(true, signature, win.RenderedText(), link.Url))
                    return false;

                PostWarning(new PendingWarning
                {
                    Hwnd = hwnd, Source = Source.Enter,
                    Url = link.Url, DisplayedText = link.TextToDisplay
                });
                return true;
            }
        }

        /// <summary>
        /// Turn the Enter key message into WM_NULL: the application then never translates it,
        /// so no WM_CHAR follows and the link is not followed.
        /// </summary>
        private static void Neutralize(IntPtr lParam)
        {
            Marshal.WriteInt32(lParam, IntPtr.Size, NativeMethods.WM_NULL);
        }

        // ---- Context menu path (repurposed HyperlinkOpen) ------------------------

        /// <summary>
        /// Called by the ribbon callback for the "Open hyperlink" menu item. Returns true
        /// to cancel Outlook's own open (a warning was queued), false to let it proceed.
        /// Never throws.
        /// </summary>
        public bool OnHyperlinkOpenRequested()
        {
            try
            {
                if (_disposed || _warningPending || _warningOpen) return _warningPending || _warningOpen;

                var settings = Settings.Load(_settings);
                string signature = settings.ExternalBannerSignature;
                if (!settings.PhishingWarnLinks || string.IsNullOrWhiteSpace(signature)) return false;

                IntPtr hwnd = IntPtr.Zero;
                if (_lastRightWwg != IntPtr.Zero && Elapsed(_lastRightTick) < RightClickMemoryMs)
                    hwnd = _lastRightWwg;
                else
                {
                    IntPtr focus = NativeMethods.GetFocus();
                    if (IsWordDocumentWindow(focus)) hwnd = focus;
                }
                if (hwnd == IntPtr.Zero) return false;

                using (var win = WordWindow.TryGet(hwnd))
                {
                    if (win == null) return false;
                    var link = win.SelectionLink();
                    if (link == null || !PhishingLinkFilter.IsWarnable(link.Url)) return false;
                    if (!PhishingWarningPolicy.ShouldWarnLink(true, signature, win.RenderedText(), link.Url))
                        return false;

                    PostWarning(new PendingWarning
                    {
                        Hwnd = hwnd, Source = Source.Menu,
                        Url = link.Url, DisplayedText = link.TextToDisplay
                    });
                    return true;
                }
            }
            catch (Exception ex)
            {
                LogHookErrorOnce(ex);
                return false; // fail open
            }
        }

        // ---- The warning ---------------------------------------------------------

        private void PostWarning(PendingWarning pending)
        {
            _warningPending = true;
            try
            {
                _dispatcher.BeginInvoke(DispatcherPriority.Normal, new Action(() => ShowWarning(pending)));
            }
            catch (Exception ex)
            {
                _warningPending = false;
                Log.Warning(ex, "Link warning: could not queue the warning.");
            }
        }

        private void ShowWarning(PendingWarning pending)
        {
            bool openAnyway;
            try
            {
                _warningOpen = true;
                var sender = _resolver.ResolveSender(pending.Hwnd);
                var message = new LinkWarningMessage(sender, pending.Url, pending.DisplayedText);
                var vm = PhishingWarningViewModel.ForLink(TaskPaneServices.Localization, message);
                openAnyway = PhishingWarningWindow.Confirm(vm);
                Log.Information("Link warning ({Source}) for {Host}: {Outcome}.",
                    pending.Source, PhishingLinkFilter.HostForLog(pending.Url),
                    openAnyway ? "opened anyway" : "cancelled");
            }
            catch (Exception ex)
            {
                // The warning itself failed: fail open and follow the link.
                Log.Error(ex, "Link warning failed; following the link.");
                openAnyway = true;
            }
            finally
            {
                _warningOpen = false;
                _warningPending = false;
            }

            if (openAnyway) FollowLink(pending);
        }

        /// <summary>
        /// Open the link after "Open anyway": re-acquire it from the Word window and call
        /// Hyperlink.Follow() (keeping Outlook's own prompts); only if it cannot be
        /// re-acquired and the scheme is http(s), hand it to the shell.
        /// </summary>
        private void FollowLink(PendingWarning pending)
        {
            try
            {
                using (var win = WordWindow.TryGet(pending.Hwnd))
                {
                    if (win != null)
                    {
                        var link = pending.Source == Source.Mouse
                            ? win.HitTest(pending.X, pending.Y)
                            : win.SelectionLink();
                        if (link != null && string.Equals(link.Url, pending.Url, StringComparison.Ordinal))
                        {
                            link.Follow();
                            return;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Log.Debug(ex, "Link warning: re-following the link through Word failed.");
            }

            try
            {
                Uri uri;
                if (Uri.TryCreate(pending.Url, UriKind.Absolute, out uri)
                    && (uri.Scheme == Uri.UriSchemeHttp || uri.Scheme == Uri.UriSchemeHttps))
                {
                    Process.Start(new ProcessStartInfo(pending.Url) { UseShellExecute = true });
                    return;
                }
                Log.Warning("Link warning: the link could not be re-acquired and its scheme is not http(s); not opened.");
            }
            catch (Exception ex)
            {
                Log.Warning(ex, "Link warning: opening the link through the shell failed.");
            }
        }

        // ---- Helpers -------------------------------------------------------------

        private static bool IsWordDocumentWindow(IntPtr hwnd)
        {
            return hwnd != IntPtr.Zero
                && string.Equals(NativeMethods.ClassNameOf(hwnd), WordDocumentWindowClass, StringComparison.Ordinal);
        }

        private static IntPtr Next(IntPtr hook, int nCode, IntPtr wParam, IntPtr lParam)
            => NativeMethods.CallNextHookEx(hook, nCode, wParam, lParam);

        private static int Elapsed(int sinceTick) => unchecked(Environment.TickCount - sinceTick);

        private void LogHookErrorOnce(Exception ex)
        {
            if (_loggedHookError) return;
            _loggedHookError = true;
            try { Log.Error(ex, "Link warning hook error (further errors are not logged)."); } catch { }
        }
    }
}
