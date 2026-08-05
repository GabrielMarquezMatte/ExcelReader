using System.Buffers;
using System.Buffers.Text;
using System.Globalization;
using System.Runtime.CompilerServices;

namespace ExcelReader.Core.Writer.Internal
{
    internal static class CellFormatter
    {
        // The 5 XML entity chars, '_' (to detect literal "_xHHHH_" escape sequences that must be
        // themselves escaped), and every C0 control char that's illegal in XML 1.0 text content
        // (0x00-0x08, 0x0B, 0x0C, 0x0E-0x1F — tab/LF/CR are legal and excluded).
        private static readonly SearchValues<char> SpecialChars = SearchValues.Create(
            "&<>\"'_" +
            "\u0000\u0001\u0002\u0003\u0004\u0005\u0006\u0007\u0008\u000B\u000C" +
            "\u000E\u000F\u0010\u0011\u0012\u0013\u0014\u0015\u0016\u0017\u0018\u0019\u001A\u001B\u001C\u001D\u001E\u001F");

        // Writes the cell reference (e.g. "B7") directly to the writer.
        // Max XLSX cell is XFD1048576 -> 3 column letters + 7 row digits.
        [SkipLocalsInit]
        private static void WriteRef(BiffBuffer xml, int columnIndex, int rowNumber)
        {
            Span<byte> buf = stackalloc byte[10];
            int len = ColumnName.Write(buf, columnIndex);
            Utf8Formatter.TryFormat(rowNumber, buf[len..], out int rowLen);
            xml.Write(buf[..(len + rowLen)]);
        }

        // Writes "<c", the optional r="..." reference, the optional s="N" style (only when styleId is
        // non-zero, so an unstyled cell costs exactly the bytes it cost before this attribute existed),
        // then typeAttr (e.g. " t=\"b\"", or empty for a plain number), then the tag's closing ">"/"/>".
        private static void WriteCellOpen(BiffBuffer xml, int columnIndex, int rowNumber, bool includeReference,
            int styleId, ReadOnlySpan<byte> typeAttr, bool selfClose)
        {
            xml.Write("<c"u8);
            if (includeReference)
            {
                xml.Write(" r=\""u8);
                WriteRef(xml, columnIndex, rowNumber);
                xml.WriteByte((byte)'"');
            }
            if (styleId != 0)
            {
                xml.Write(" s=\""u8);
                WriteValue(xml, styleId, sizeHint: 8);
                xml.WriteByte((byte)'"');
            }
            xml.Write(typeAttr);
            xml.Write(selfClose ? "/>"u8 : ">"u8);
        }

        internal static void WriteEmpty(BiffBuffer xml, int columnIndex, int rowNumber, bool includeReference, int styleId = 0)
        {
            WriteCellOpen(xml, columnIndex, rowNumber, includeReference, styleId, default, selfClose: true);
        }

        internal static void WriteString(BiffBuffer xml, string value, int columnIndex, int rowNumber, bool includeReference, int styleId = 0)
        {
            WriteCellOpen(xml, columnIndex, rowNumber, includeReference, styleId, " t=\"inlineStr\""u8, selfClose: false);
            xml.Write(HasEdgeWhitespace(value) ? "<is><t xml:space=\"preserve\">"u8 : "<is><t>"u8);
            WriteEscaped(xml, value);
            xml.Write("</t></is></c>"u8);
        }

        private static bool HasEdgeWhitespace(string value)
        {
            return value.Length != 0 && (char.IsWhiteSpace(value[0]) || char.IsWhiteSpace(value[^1]));
        }

        internal static void WriteSharedString(BiffBuffer xml, int sharedStringIndex, int columnIndex, int rowNumber, bool includeReference, int styleId = 0)
        {
            WriteCellOpen(xml, columnIndex, rowNumber, includeReference, styleId, " t=\"s\""u8, selfClose: false);
            xml.Write("<v>"u8);
            WriteValue(xml, sharedStringIndex, sizeHint: 16);
            xml.Write("</v></c>"u8);
        }

        internal static void WriteBool(BiffBuffer xml, bool value, int columnIndex, int rowNumber, bool includeReference, int styleId = 0)
        {
            WriteCellOpen(xml, columnIndex, rowNumber, includeReference, styleId, " t=\"b\""u8, selfClose: false);
            xml.Write("<v>"u8);
            xml.WriteByte(value ? (byte)'1' : (byte)'0');
            xml.Write("</v></c>"u8);
        }

        // styleId has no zero-means-omit default here: a date cell always needs an explicit style (it
        // is how the serial number renders as a date instead of a plain number), so the caller always
        // passes either the builtin date style (1) or an active row/column style override.
        internal static void WriteDateTime(BiffBuffer xml, DateTime value, int columnIndex, int rowNumber, bool includeReference, int styleId)
        {
            WriteCellOpen(xml, columnIndex, rowNumber, includeReference, styleId, default, selfClose: false);
            xml.Write("<v>"u8);
            WriteValue(xml, ExcelEpoch.OADateToSerial(value.ToOADate(), date1904: false), sizeHint: 32);
            xml.Write("</v></c>"u8);
        }

        internal static void WriteNumber<T>(BiffBuffer xml, T value, int columnIndex, int rowNumber, bool includeReference, int styleId = 0)
            where T : IUtf8SpanFormattable
        {
            WriteCellOpen(xml, columnIndex, rowNumber, includeReference, styleId, default, selfClose: false);
            xml.Write("<v>"u8);
            WriteValue(xml, value, sizeHint: 64);
            xml.Write("</v></c>"u8);
        }

        internal static void WriteNumber(BiffBuffer xml, int value, int columnIndex, int rowNumber, bool includeReference, int styleId = 0)
        {
            WriteCellOpen(xml, columnIndex, rowNumber, includeReference, styleId, default, selfClose: false);
            xml.Write("<v>"u8);
            WriteValue(xml, value, sizeHint: 16);
            xml.Write("</v></c>"u8);
        }

        internal static void WriteNumber(BiffBuffer xml, long value, int columnIndex, int rowNumber, bool includeReference, int styleId = 0)
        {
            WriteCellOpen(xml, columnIndex, rowNumber, includeReference, styleId, default, selfClose: false);
            xml.Write("<v>"u8);
            WriteValue(xml, value, sizeHint: 32);
            xml.Write("</v></c>"u8);
        }

        internal static void WriteNumber(BiffBuffer xml, double value, int columnIndex, int rowNumber, bool includeReference, int styleId = 0)
        {
            WriteCellOpen(xml, columnIndex, rowNumber, includeReference, styleId, default, selfClose: false);
            xml.Write("<v>"u8);
            WriteValue(xml, value, sizeHint: 32);
            xml.Write("</v></c>"u8);
        }

        internal static void WriteNumber(BiffBuffer xml, decimal value, int columnIndex, int rowNumber, bool includeReference, int styleId = 0)
        {
            WriteCellOpen(xml, columnIndex, rowNumber, includeReference, styleId, default, selfClose: false);
            xml.Write("<v>"u8);
            WriteValue(xml, value, sizeHint: 64);
            xml.Write("</v></c>"u8);
        }

        // Utf8Formatter is culture-free (no NumberFormatInfo.GetInstance lookup per cell) and matches
        // IUtf8SpanFormattable's default/InvariantCulture output for these types exactly, including
        // double's shortest-round-trip format since .NET Core 3.0. These non-generic overloads bind
        // ahead of the generic one below for int/long/double call sites.
        private static void WriteValue(BiffBuffer xml, int value, int sizeHint)
        {
            int size = sizeHint;
            int written;
            while (!Utf8Formatter.TryFormat(value, xml.GetSpan(size), out written))
            {
                size = checked(size * 2);
            }
            xml.Advance(written);
        }

        private static void WriteValue(BiffBuffer xml, long value, int sizeHint)
        {
            int size = sizeHint;
            int written;
            while (!Utf8Formatter.TryFormat(value, xml.GetSpan(size), out written))
            {
                size = checked(size * 2);
            }
            xml.Advance(written);
        }

        private static void WriteValue(BiffBuffer xml, double value, int sizeHint)
        {
            CellValueGuards.ThrowIfNonFinite(value, nameof(value));
            int size = sizeHint;
            int written;
            while (!Utf8Formatter.TryFormat(value, xml.GetSpan(size), out written))
            {
                size = checked(size * 2);
            }
            xml.Advance(written);
        }

        // Formats a numeric value straight into the buffer's free tail (no temp span + copy). The default
        // format is shortest round-trippable for floating point, so cells stay small and exactly readable.
        // Used by decimal and the generic WriteNumber<T> overload (Utf8Formatter doesn't cover either).
        private static void WriteValue<T>(BiffBuffer xml, T value, int sizeHint)
            where T : IUtf8SpanFormattable
        {
            if (typeof(T) == typeof(double))
            {
                CellValueGuards.ThrowIfNonFinite(Unsafe.As<T, double>(ref value), nameof(value));
            }
            else if (typeof(T) == typeof(float))
            {
                CellValueGuards.ThrowIfNonFinite(Unsafe.As<T, float>(ref value), nameof(value));
            }
            int size = sizeHint;
            int written;
            while (!value.TryFormat(xml.GetSpan(size), out written, default, CultureInfo.InvariantCulture))
            {
                size = checked(size * 2);
            }
            xml.Advance(written);
        }

        internal static void WriteEscaped(BiffBuffer xml, ReadOnlySpan<char> value)
        {
            int start = 0;
            int next = value.IndexOfAny(SpecialChars);
            while (next >= 0)
            {
                int i = start + next;
                char c = value[i];
                if (TryGetEntity(c, out ReadOnlySpan<byte> entity))
                {
                    if (i > start)
                    {
                        xml.WriteUtf8(value[start..i]);
                    }
                    xml.Write(entity);
                    start = i + 1;
                }
                else if (c == '_')
                {
                    // A literal "_xHHHH_" in the source text must itself be escaped, or Excel reads it
                    // back as a ST_Xstring unicode escape instead of the literal characters the writer
                    // put there (the underscore's own escape is "_x005F_").
                    if (IsXHHHHUnderscorePattern(value, i))
                    {
                        if (i > start)
                        {
                            xml.WriteUtf8(value[start..i]);
                        }
                        xml.Write("_x005F_"u8);
                        start = i + 1;
                    }
                    else
                    {
                        // A plain '_' remains in the pending run, but the next scan must move past
                        // it; otherwise this loop would rediscover the same underscore forever.
                        int following = value[(i + 1)..].IndexOfAny(SpecialChars);
                        if (following < 0)
                        {
                            break;
                        }
                        next = following + (i + 1 - start);
                        continue;
                    }
                }
                else
                {
                    // Illegal XML 1.0 control character: encode as ST_Xstring's "_xHHHH_" escape, which
                    // Excel writes and reads for exactly this case, instead of emitting invalid XML.
                    if (i > start)
                    {
                        xml.WriteUtf8(value[start..i]);
                    }
                    WriteHexEscape(xml, c);
                    start = i + 1;
                }
                next = value[start..].IndexOfAny(SpecialChars);
            }
            if (start < value.Length)
            {
                xml.WriteUtf8(value[start..]);
            }
        }

        private static bool TryGetEntity(char c, out ReadOnlySpan<byte> entity)
        {
            switch (c)
            {
                case '&': entity = "&amp;"u8; return true;
                case '<': entity = "&lt;"u8; return true;
                case '>': entity = "&gt;"u8; return true;
                case '"': entity = "&quot;"u8; return true;
                case '\'': entity = "&apos;"u8; return true;
                default: entity = default; return false;
            }
        }

        // True when value[i..] starts with the ECMA-376 ST_Xstring escape shape "_xHHHH_" (4 hex digits).
        private static bool IsXHHHHUnderscorePattern(ReadOnlySpan<char> value, int i)
        {
            if (i + 6 >= value.Length || (value[i + 1] != 'x' && value[i + 1] != 'X'))
            {
                return false;
            }
            for (int k = 0; k < 4; k++)
            {
                if (!Uri.IsHexDigit(value[i + 2 + k]))
                {
                    return false;
                }
            }
            return value[i + 6] == '_';
        }

        [SkipLocalsInit]
        private static void WriteHexEscape(BiffBuffer xml, char c)
        {
            Span<byte> buf = stackalloc byte[7];
            "_x0000_"u8.CopyTo(buf);
            int code = c;
            buf[2] = HexDigit((code >> 12) & 0xF);
            buf[3] = HexDigit((code >> 8) & 0xF);
            buf[4] = HexDigit((code >> 4) & 0xF);
            buf[5] = HexDigit(code & 0xF);
            xml.Write(buf);
        }

        private static byte HexDigit(int nibble)
        {
            return (byte)(nibble < 10 ? '0' + nibble : 'A' + (nibble - 10));
        }
    }
}
