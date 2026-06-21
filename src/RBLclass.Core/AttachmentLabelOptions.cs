using System;

namespace RBLclass.Core
{
    /// <summary>
    /// Localized templates the shell supplies to <see cref="AttachmentLabelFormatter"/>
    /// so the label composer stays UI-free (v2.4.0.0 F3).
    /// </summary>
    public sealed class AttachmentLabelOptions
    {
        public AttachmentLabelOptions(string headerOne, string headerOther,
                                      string savedToTemplate, string deletedOnTemplate,
                                      string dateFormat)
        {
            HeaderOne = headerOne ?? string.Empty;
            HeaderOther = headerOther ?? string.Empty;
            SavedToTemplate = savedToTemplate ?? "{0}";
            DeletedOnTemplate = deletedOnTemplate ?? "{0}";
            DateFormat = string.IsNullOrEmpty(dateFormat) ? "yyyy-MM-dd" : dateFormat;
        }

        /// <summary>Heading when exactly one attachment was disposed of.</summary>
        public string HeaderOne { get; }

        /// <summary>Heading when two or more attachments were disposed of.</summary>
        public string HeaderOther { get; }

        /// <summary>Template for a saved attachment, e.g. "Saved to {0}" ({0} = directory).</summary>
        public string SavedToTemplate { get; }

        /// <summary>Template for a deleted attachment, e.g. "Deleted on {0}" ({0} = date).</summary>
        public string DeletedOnTemplate { get; }

        /// <summary>Date format for the deleted-on date.</summary>
        public string DateFormat { get; }
    }
}
