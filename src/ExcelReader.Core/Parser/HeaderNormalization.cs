using System.Globalization;
using System.Text;

namespace ExcelReader.Core.Parser
{
    [Flags]
    public enum HeaderNormalization
    {
        None = 0,
        Trim = 1 << 0,
        CollapseSpaces = 1 << 1,
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
