using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using RBLclass.AddIn.Localization;
using RBLclass.AddIn.ViewModels;
using RBLclass.AddIn.Views;
using RBLclass.Core;
using RBLclass.Core.Persistence;
using Serilog;
using OutlookOM = Microsoft.Office.Interop.Outlook;

namespace RBLclass.AddIn.Phishing
{
    /// <summary>
    /// Warns before an attachment from a mail carrying the organization's external-sender
    /// banner opens (v2.13.0.0 D2). It hooks the item events BeforeAttachmentPreview and
    /// BeforeAttachmentWriteToTempFile (both can cancel) on the single selected mail of
    /// the explorer and on the item of each opened inspector.
    ///
    /// Rules: the item and its delegates are kept in fields so the events keep firing;
    /// every hook is released on a change, on Inspector.Close and on Dispose; the mail
    /// body is read only when an event fires; every handler fails open (an error logs
    /// and lets the attachment open); one warning at a time, and a short grace period
    /// after "Open anyway" so Outlook's follow-up events for the same action do not
    /// warn twice. Everything runs on the Outlook UI thread. Only counts and outcomes
    /// are logged, never file names or addresses.
    /// </summary>
    internal sealed class AttachmentWarningGuard : IDisposable
    {
        private const int GraceMilliseconds = 4000;

        private sealed class Hooked
        {
            public OutlookOM.MailItem Mail;
            public string EntryId;
            public OutlookOM.Inspector Inspector;
            public OutlookOM.ItemEvents_10_BeforeAttachmentPreviewEventHandler OnPreview;
            public OutlookOM.ItemEvents_10_BeforeAttachmentWriteToTempFileEventHandler OnWrite;
            public OutlookOM.InspectorEvents_10_CloseEventHandler OnClose;
        }

        private readonly IMailStore _store;
        private readonly ISettingsStore _settings;
        private readonly List<Hooked> _inspectors = new List<Hooked>();
        private Hooked _selected;
        private bool _dialogOpen;
        private bool _disposed;
        private int _graceUntil;
        private string _graceEntryId;

        public AttachmentWarningGuard(IMailStore store, ISettingsStore settings)
        {
            _store = store;
            _settings = settings;
        }

        /// <summary>
        /// Hook the single selected mail (replacing the previous selection hook). Takes
        /// ownership of <paramref name="mail"/>: the caller must not release it.
        /// </summary>
        public void HookSelected(OutlookOM.MailItem mail)
        {
            if (_disposed || mail == null) { Release(mail); return; }

            string entryId = SafeEntryId(mail);
            if (_selected != null && entryId != null && entryId == _selected.EntryId)
            {
                Release(mail); // same mail: keep the existing hook
                return;
            }

            UnhookSelected();
            _selected = Attach(mail, entryId, null);
        }

        public void UnhookSelected()
        {
            var h = _selected;
            _selected = null;
            Detach(h);
        }

        /// <summary>
        /// Hook the item of an opened inspector; released on Inspector.Close. NewInspector
        /// fires before Outlook has the item ready (its CurrentItem is not yet a MailItem),
        /// so the hook is made on the first Inspector.Activate, like the reply banner strip.
        /// </summary>
        public void HookInspector(OutlookOM.Inspector inspector)
        {
            if (_disposed || inspector == null) return;

            var events = inspector as OutlookOM.InspectorEvents_10_Event;
            if (events == null) return;

            OutlookOM.InspectorEvents_10_ActivateEventHandler onActivate = null;
            onActivate = () =>
            {
                try
                {
                    events.Activate -= onActivate;
                    HookInspectorNow(inspector);
                }
                catch (Exception ex) { Log.Debug(ex, "Attachment guard: hooking on Inspector.Activate failed."); }
            };
            events.Activate += onActivate;
        }

        private void HookInspectorNow(OutlookOM.Inspector inspector)
        {
            if (_disposed || inspector == null) return;

            object item = null;
            try { item = inspector.CurrentItem; }
            catch (Exception ex) { Log.Debug(ex, "Attachment guard: Inspector.CurrentItem failed."); }
            var mail = item as OutlookOM.MailItem;
            if (mail == null)
            {
                string cls = "?";
                try { cls = Convert.ToString(item.GetType().InvokeMember("Class",
                    System.Reflection.BindingFlags.GetProperty, null, item, null)); } catch { }
                Log.Debug("Attachment guard: inspector item is not a MailItem (Class={Class}, _MailItem={Underscore}).",
                    cls, item is OutlookOM._MailItem);
                Release(item);
                return;
            }

            var hooked = Attach(mail, SafeEntryId(mail), inspector);
            if (hooked == null) return;

            var inspectorEvents = inspector as OutlookOM.InspectorEvents_10_Event;
            if (inspectorEvents != null)
            {
                hooked.OnClose = () =>
                {
                    try
                    {
                        _inspectors.Remove(hooked);
                        Detach(hooked);
                    }
                    catch (Exception ex) { Log.Debug(ex, "Attachment guard: inspector close cleanup failed."); }
                };
                try { inspectorEvents.Close += hooked.OnClose; }
                catch (Exception ex) { Log.Debug(ex, "Attachment guard: Inspector.Close subscribe failed."); }
            }
            _inspectors.Add(hooked);
        }

        public void Dispose()
        {
            if (_disposed) return;
            _disposed = true;
            UnhookSelected();
            foreach (var h in _inspectors.ToArray()) Detach(h);
            _inspectors.Clear();
        }

        private Hooked Attach(OutlookOM.MailItem mail, string entryId, OutlookOM.Inspector inspector)
        {
            var events = mail as OutlookOM.ItemEvents_10_Event;
            if (events == null)
            {
                Log.Debug("Attachment guard: the mail does not expose the item events.");
                Release(mail);
                return null;
            }

            var hooked = new Hooked { Mail = mail, EntryId = entryId, Inspector = inspector };
            hooked.OnPreview = (OutlookOM.Attachment attachment, ref bool cancel) =>
            {
                if (!Evaluate(hooked, attachment, "preview")) cancel = true;
            };
            hooked.OnWrite = (OutlookOM.Attachment attachment, ref bool cancel) =>
            {
                if (!Evaluate(hooked, attachment, "open")) cancel = true;
            };

            try
            {
                events.BeforeAttachmentPreview += hooked.OnPreview;
                events.BeforeAttachmentWriteToTempFile += hooked.OnWrite;
            }
            catch (Exception ex)
            {
                Log.Warning(ex, "Attachment guard: hooking the item events failed.");
                Detach(hooked);
                return null;
            }
            Log.Debug("Attachment guard: hooked a mail ({Kind}).", inspector != null ? "inspector" : "selection");
            return hooked;
        }

        private void Detach(Hooked h)
        {
            if (h == null) return;

            var events = h.Mail as OutlookOM.ItemEvents_10_Event;
            if (events != null)
            {
                try { if (h.OnPreview != null) events.BeforeAttachmentPreview -= h.OnPreview; } catch { }
                try { if (h.OnWrite != null) events.BeforeAttachmentWriteToTempFile -= h.OnWrite; } catch { }
            }
            if (h.Inspector != null && h.OnClose != null)
            {
                var inspectorEvents = h.Inspector as OutlookOM.InspectorEvents_10_Event;
                try { if (inspectorEvents != null) inspectorEvents.Close -= h.OnClose; } catch { }
            }
            Release(h.Mail);
            h.Mail = null;
            h.Inspector = null;
        }

        /// <summary>Returns true to let the attachment open, false to cancel it.</summary>
        private bool Evaluate(Hooked hooked, OutlookOM.Attachment attachment, string trigger)
        {
            try
            {
                Log.Debug("Attachment guard: {Trigger} event on a {Kind} mail.", trigger, hooked.Inspector != null ? "inspector" : "selection");
                var mail = hooked.Mail;
                if (_disposed || mail == null) return true;

                // The first warning is on screen: this is a duplicate event of the same action.
                if (_dialogOpen) return false;

                // Right after "Open anyway": Outlook's follow-up events for the same mail.
                if (_graceEntryId != null && hooked.EntryId == _graceEntryId
                    && unchecked(Environment.TickCount - _graceUntil) < 0)
                    return true;

                var settings = Settings.Load(_settings);
                string signature = settings.ExternalBannerSignature;
                if (!settings.PhishingWarnAttachments || string.IsNullOrWhiteSpace(signature))
                    return true;

                string html = _store.GetLiveItemHtmlBody(mail);
                if (!PhishingWarningPolicy.ShouldWarnAttachment(true, signature, html))
                    return true;

                string fileName = null;
                if (attachment != null)
                {
                    try { fileName = attachment.FileName; } catch { }
                }
                var message = new PhishingWarningMessage(_store.GetSenderInfo(mail), fileName);

                _dialogOpen = true;
                try
                {
                    var vm = PhishingWarningViewModel.ForAttachment(TaskPaneServices.Localization, message);
                    bool openAnyway = PhishingWarningWindow.Confirm(vm);
                    Log.Information("Attachment warning ({Trigger}): {Outcome}.",
                        trigger, openAnyway ? "opened anyway" : "cancelled");
                    if (openAnyway)
                    {
                        _graceEntryId = hooked.EntryId;
                        _graceUntil = unchecked(Environment.TickCount + GraceMilliseconds);
                    }
                    return openAnyway;
                }
                finally { _dialogOpen = false; }
            }
            catch (Exception ex)
            {
                // Fail open, like the send guard: a defect here must never block the attachment.
                Log.Error(ex, "Attachment warning failed; letting the attachment open.");
                return true;
            }
        }

        private static string SafeEntryId(OutlookOM.MailItem mail)
        {
            try { return mail.EntryID; } catch { return null; }
        }

        private static void Release(object com)
        {
            if (com == null) return;
            try { if (Marshal.IsComObject(com)) Marshal.ReleaseComObject(com); } catch { }
        }
    }
}
