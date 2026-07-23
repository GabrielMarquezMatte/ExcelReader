using System.Globalization;
using System.Text;

namespace ExcelReader.Core.Parser
{
    /// <summary>
    /// Flags controlling how header text is normalized before being matched against
    /// <see cref="ExcelColumnAttribute.Name"/> or a property name.
    /// </summary>
    [Flags]
    public enum HeaderNormalization
    {
        /// <summary>Apply no normalization; header text is matched as-is.</summary>
        None = 0,
        /// <summary>Trim leading and trailing whitespace.</summary>
        Trim = 1 << 0,
        /// <summary>Collapse runs of whitespace into a single space.</summary>
        CollapseSpaces = 1 << 1,
        /// <summary>Strip diacritical marks (e.g. accents) from characters.</summary>
        RemoveDiacritics = 1 << 2,
    }

    internal static class HeaderNormalizationExtensions
    {
        internal static string Apply(this HeaderNormalization norm, string value)
        {
            if (norm == HeaderNormalization.None)
            {
                return value;
            }
            if (norm.HasFlag(HeaderNormalization.Trim))
            {
                value = value.Trim();
            }
            if (norm.HasFlag(HeaderNormalization.CollapseSpaces))
            {
                value = string.Join(' ', value.Split([' ', '\t', '\r', '\n'], StringSplitOptions.RemoveEmptyEntries));
            }
            if (norm.HasFlag(HeaderNormalization.RemoveDiacritics))
            {
                string nfd = value.Normalize(NormalizationForm.FormD);
                value = new string([.. nfd.Where(c => CharUnicodeInfo.GetUnicodeCategory(c) != UnicodeCategory.NonSpacingMark)]);
            }
            return value;
        }
    }
}
