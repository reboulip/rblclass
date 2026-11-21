using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Threading;
using Serilog;

namespace RBLclass.Outlook.Adapter
{
    /// <summary>One MHTML to PDF conversion.</summary>
    public sealed class PdfJob
    {
        public PdfJob(string mhtPath, string pdfPath)
        {
            MhtPath = mhtPath;
            PdfPath = pdfPath;
        }

        public string MhtPath { get; }
        public string PdfPath { get; }
    }

    /// <summary>Outcome of a batch conversion.</summary>
    public sealed class PdfConversionResult
    {
        public PdfConversionResult(int converted, int failed, bool wordUnavailable)
        {
            Converted = converted;
            Failed = failed;
            WordUnavailable = wordUnavailable;
        }

        public int Converted { get; }
        public int Failed { get; }

        /// <summary>True when Word is missing or could not start: nothing was written.</summary>
        public bool WordUnavailable { get; }
    }

    /// <summary>
    /// Converts MHTML files to PDF through Microsoft Word automation (v2.13.0.0 C1).
    /// Word is reached by late binding only: no Word interop reference, no new
    /// dependency, optional at runtime. Word runs out of process, on a dedicated STA
    /// thread; this class never touches Outlook objects (Outlook OM calls stay on the
    /// UI thread). One Word instance per batch, always quit and released.
    /// </summary>
    public sealed class WordPdfConverter
    {
        private const int WdExportFormatPdf = 17;

        /// <summary>True when a "Word.Application" COM class is registered.</summary>
        public static bool IsWordInstalled()
        {
            try { return Type.GetTypeFromProgID("Word.Application", false) != null; }
            catch { return false; }
        }

        /// <summary>
        /// Convert every job on a dedicated STA thread and wait for it. Blocking, so
        /// call it from a worker (<c>Task.Run</c>), never from the Outlook UI thread.
        /// <paramref name="progress"/> is invoked on the worker thread with
        /// (completed, total) after each file.
        /// </summary>
        public PdfConversionResult ConvertAll(IReadOnlyList<PdfJob> jobs, Action<int, int> progress)
        {
            PdfConversionResult result = new PdfConversionResult(0, jobs.Count, wordUnavailable: false);
            var thread = new Thread(() => { result = ConvertOnThisThread(jobs, progress); })
            {
                IsBackground = true,
                Name = "RBLclass PDF export"
            };
            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();
            thread.Join();
            return result;
        }

        private static PdfConversionResult ConvertOnThisThread(IReadOnlyList<PdfJob> jobs, Action<int, int> progress)
        {
            int converted = 0, failed = 0;
            object word = null;
            try
            {
                var type = Type.GetTypeFromProgID("Word.Application", false);
                if (type == null) return new PdfConversionResult(0, 0, wordUnavailable: true);

                try
                {
                    word = Activator.CreateInstance(type);
                    ((dynamic)word).Visible = false;
                    ((dynamic)word).DisplayAlerts = 0; // wdAlertsNone
                }
                catch (Exception ex)
                {
                    Log.Warning(ex, "PDF export: Word could not be started.");
                    return new PdfConversionResult(0, 0, wordUnavailable: true);
                }

                dynamic app = word;
                for (int i = 0; i < jobs.Count; i++)
                {
                    object doc = null;
                    try
                    {
                        // Documents.Open(FileName, ConfirmConversions, ReadOnly, AddToRecentFiles,
                        // PasswordDocument, PasswordTemplate, Revert, WritePasswordDocument,
                        // WritePasswordTemplate, Format, Encoding, Visible)
                        doc = app.Documents.Open(jobs[i].MhtPath, false, true, false,
                                                 Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                 Type.Missing, Type.Missing, Type.Missing, false);
                        ((dynamic)doc).ExportAsFixedFormat(jobs[i].PdfPath, WdExportFormatPdf);
                        converted++;
                    }
                    catch (Exception ex)
                    {
                        failed++;
                        Log.Error(ex, "PDF export: conversion failed for item {Index}.", i);
                    }
                    finally
                    {
                        if (doc != null)
                        {
                            try { ((dynamic)doc).Close(0); } catch { }
                            try { Marshal.ReleaseComObject(doc); } catch { }
                        }
                    }

                    try { progress?.Invoke(i + 1, jobs.Count); } catch { }
                }
            }
            finally
            {
                if (word != null)
                {
                    try { ((dynamic)word).Quit(0); } catch { }
                    try { Marshal.ReleaseComObject(word); } catch { }
                }
            }
            return new PdfConversionResult(converted, failed, wordUnavailable: false);
        }
    }
}
