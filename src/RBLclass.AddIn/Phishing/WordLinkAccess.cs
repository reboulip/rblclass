using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;

namespace RBLclass.AddIn.Phishing
{
    /// <summary>A hyperlink found in a Word document window.</summary>
    internal sealed class WordLink
    {
        public WordLink(string url, string textToDisplay, object hyperlink)
        {
            Url = url ?? string.Empty;
            TextToDisplay = textToDisplay ?? string.Empty;
            Hyperlink = hyperlink;
        }

        public string Url { get; }
        public string TextToDisplay { get; }

        /// <summary>The Word Hyperlink COM object; owned (and released) by its <see cref="WordWindow"/>.</summary>
        public object Hyperlink { get; }

        /// <summary>Follow the link the way Outlook would (keeps Outlook's own security prompts).</summary>
        public void Follow() => ((dynamic)Hyperlink).Follow();
    }

    /// <summary>
    /// Late-bound access to the Word document behind a <c>_WwG</c> window (v2.13.0.0 D3):
    /// the window's native object model is obtained through
    /// <c>AccessibleObjectFromWindow(OBJID_NATIVEOM)</c> and driven through
    /// <c>dynamic</c> (IDispatch), so there is no Word interop reference. Every COM
    /// object obtained is tracked and released on <see cref="Dispose"/>. Runs on the
    /// Outlook UI thread (in-process, same STA).
    /// </summary>
    internal sealed class WordWindow : IDisposable
    {
        private readonly object _window;
        private readonly List<object> _owned = new List<object>();

        private WordWindow(object window)
        {
            _window = window;
            _owned.Add(window);
        }

        /// <summary>The window's object model, or null when the window does not expose it.</summary>
        public static WordWindow TryGet(IntPtr hwnd)
        {
            try
            {
                var iid = NativeMethods.IID_IDispatch;
                object obj;
                int hr = NativeMethods.AccessibleObjectFromWindow(hwnd, NativeMethods.OBJID_NATIVEOM, ref iid, out obj);
                return hr == 0 && obj != null ? new WordWindow(obj) : null;
            }
            catch
            {
                return null;
            }
        }

        private object Track(object com)
        {
            if (com != null && Marshal.IsComObject(com)) _owned.Add(com);
            return com;
        }

        /// <summary>True for a protected (read-only) document: a received mail in a reading pane or window.</summary>
        public bool IsReadOnly()
        {
            dynamic w = _window;
            dynamic doc = Track((object)w.Document);
            return (int)doc.ProtectionType == 3; // wdAllowOnlyReading
        }

        /// <summary>Whether a link needs Ctrl+click to open (the reply editor's default).</summary>
        public bool CtrlClickRequired()
        {
            dynamic w = _window;
            dynamic app = Track((object)w.Application);
            dynamic options = Track((object)app.Options);
            return (bool)options.CtrlClickHyperlinkToOpen;
        }

        /// <summary>The document's rendered text (what the user reads, whatever the stored encoding).</summary>
        public string RenderedText()
        {
            dynamic w = _window;
            dynamic doc = Track((object)w.Document);
            dynamic content = Track((object)doc.Content);
            return (string)content.Text;
        }

        /// <summary>
        /// The link under a screen point, or null. A point strictly inside a link wins;
        /// otherwise the link that starts at the nearest character counts when the point
        /// is inside the link's own box (image links, the left edge of a link).
        /// </summary>
        public WordLink HitTest(int x, int y)
        {
            dynamic w = _window;
            dynamic p = Track((object)w.RangeFromPoint(x, y));
            if (p == null) return null;

            dynamic pLinks = Track((object)p.Hyperlinks);
            if ((int)pLinks.Count > 0)
                return Make(Track((object)pLinks.Item(1)));

            dynamic doc = Track((object)w.Document);
            dynamic content = Track((object)doc.Content);
            int start = (int)p.Start;
            int end = (int)content.End;
            dynamic r1 = Track((object)doc.Range(start, Math.Min(start + 1, end)));
            dynamic r1Links = Track((object)r1.Hyperlinks);
            if ((int)r1Links.Count == 0) return null;

            dynamic hl = Track((object)r1Links.Item(1));
            dynamic hlRange = Track((object)hl.Range);
            int l = 0, t = 0, width = 0, height = 0;
            w.GetPoint(ref l, ref t, ref width, ref height, hlRange);
            if (x >= l - 2 && x <= l + width + 2 && y >= t - 2 && y <= t + height + 2)
                return Make(hl);
            return null;
        }

        /// <summary>The link at the current selection (Enter and context-menu paths), or null.</summary>
        public WordLink SelectionLink()
        {
            dynamic w = _window;
            dynamic selection = Track((object)w.Selection);
            if (selection == null) return null;
            dynamic links = Track((object)selection.Hyperlinks);
            if ((int)links.Count == 0) return null;
            return Make(Track((object)links.Item(1)));
        }

        private WordLink Make(object hyperlink)
        {
            dynamic hl = hyperlink;
            string address = (string)hl.Address ?? string.Empty;
            string sub = (string)hl.SubAddress ?? string.Empty;
            string url = sub.Length > 0 ? address + "#" + sub : address;
            string shown = string.Empty;
            try { shown = (string)hl.TextToDisplay ?? string.Empty; } catch { }
            return new WordLink(url, shown, hyperlink);
        }

        public void Dispose()
        {
            for (int i = _owned.Count - 1; i >= 0; i--)
            {
                try { Marshal.ReleaseComObject(_owned[i]); } catch { }
            }
            _owned.Clear();
        }
    }
}
