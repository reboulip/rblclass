using System.Globalization;
using System.Text;

namespace RBLclass.Core
{
    /// <summary>
    /// Shared text normalization for folder search: case-folding + diacritic
    /// removal, so "Règlement" and "reglement" compare equal (the legacy tool
    /// used vbTextCompare + space stripping; we additionally fold accents).
    /// </summary>
    public static class TextNormalization
    {
        /// <summary>
        /// Lower-case (invariant) with diacritics removed. Returns empty for
        /// null/empty input.
        /// </summary>
        public static string Normalize(string value)
        {
            if (string.IsNullOrEmpty(value))
                return string.Empty;

            string decomposed = value.Normalize(NormalizationForm.FormD);
            var sb = new StringBuilder(decomposed.Length);
            foreach (char ch in decomposed)
            {
                if (CharUnicodeInfo.GetUnicodeCategory(ch) != UnicodeCategory.NonSpacingMark)
                    sb.Append(ch);
            }

            return sb.ToString()
                     .Normalize(NormalizationForm.FormC)
                     .ToLowerInvariant();
        }
    }
}
