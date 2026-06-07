using System;
using System.Collections.Generic;
using System.Linq;

namespace RBLclass.Core
{
    /// <summary>
    /// Legacy 6b: warn before sending a mail with no attachment whose body
    /// mentions one. ⚠ DEVIATION (robustness): the legacy isolates "the
    /// latest reply" via the theme-specific HTML separator
    /// "solid #B5C4DF 1.0pt" before scanning - brittle across Outlook
    /// versions/themes/locales. We scan the whole body instead: simpler,
    /// portable, and the same de-risking trade as Step 6's conversation
    /// widening (a small chance of a false positive from quoted text below
    /// versus a heuristic that silently stops matching one day).
    /// </summary>
    public sealed class ForgottenAttachmentGuard
    {
        public bool ShouldWarn(string bodyText, int attachmentCount, IReadOnlyList<string> keywords)
        {
            if (attachmentCount > 0) return false;
            if (string.IsNullOrEmpty(bodyText) || keywords == null) return false;

            return keywords.Any(keyword =>
                !string.IsNullOrEmpty(keyword) &&
                bodyText.IndexOf(keyword, StringComparison.OrdinalIgnoreCase) >= 0);
        }
    }
}
