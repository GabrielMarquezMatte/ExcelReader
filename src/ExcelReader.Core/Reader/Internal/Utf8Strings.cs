using System.Text;

namespace ExcelReader.Core.Reader.Internal
{
    internal static class Utf8Strings
    {
        /// <summary>Materializes UTF-8 cell text. ASCII widens straight into the new string; Encoding.UTF8
        /// counts characters in one pass and transcodes in a second, which dominates short cell values.</summary>
        internal static string Create(ReadOnlySpan<byte> utf8)
        {
            if (!Ascii.IsValid(utf8))
            {
                return Encoding.UTF8.GetString(utf8);
            }
            return string.Create(utf8.Length, utf8, static (chars, ascii) => Ascii.ToUtf16(ascii, chars, out _));
        }
    }
}
