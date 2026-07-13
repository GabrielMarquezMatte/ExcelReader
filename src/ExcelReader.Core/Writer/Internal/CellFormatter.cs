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
        private static readonly SearchValues<char> specialChars = SearchValues.Create(
            "&<>\"'_" +
            "\u0000\u0001\u0002\u0003\u0004\u0005\u0006\u0007\u0008\u000B\u000C" +
            "\u000E\u000F\u0010\u0011\u0012\u0013\u0014\u0015\u0016\u0017\u0018\u0019\u001A\u001B\u001C\u001D\u001E\u001F");

        // Writes the cell reference (e.g. "B7") directly to the writer.
        // Max XLSX cell is XFD1048576 -> 3 column letters + 7 row digits.
        private static void WriteRef(BiffBuffer xml, int columnIndex, int rowNumber)
        {
            Span<byte> buf = stackalloc byte[10];
            int len = ColumnName.Write(buf, columnIndex);
            Utf8Formatter.TryFormat(rowNumber, buf[len..], out int rowLen);
            xml.Write(buf[..(len + rowLen)]);
        }

        private static void WriteCellOpenWithRef(BiffBuffer xml, int columnIndex, int rowNumber, ReadOnlySpan<byte> tail)
        {
            xml.Write("<c r=\""u8);
            WriteRef(xml, columnIndex, rowNumber);
            xml.WriteByte((byte)'"');
            xml.Write(tail);
        }

        internal static void WriteEmpty(BiffBuffer xml, int columnIndex, int rowNumber, bool includeReference)
        {
            if (includeReference)
            {
                WriteCellOpenWithRef(xml, columnIndex, rowNumber, "/>"u8);
                return;
            }
            xml.Write("<c/>"u8);
        }

        internal static void WriteString(BiffBuffer xml, string value, int columnIndex, int rowNumber, bool includeReference)
        {
            if (includeReference)
            {
                WriteCellOpenWithRef(xml, columnIndex, rowNumber, HasEdgeWhitespace(value)
                    ? " t=\"inlineStr\"><is><t xml:space=\"preserve\">"u8
                    : " t=\"inlineStr\"><is><t>"u8);
            }
            else
            {
                xml.Write(HasEdgeWhitespace(value)
                    ? "<c t=\"inlineStr\"><is><t xml:space=\"preserve\">"u8
                    : "<c t=\"inlineStr\"><is><t>"u8);
            }
            WriteEscaped(xml, value);
            xml.Write("</t></is></c>"u8);
        }

        internal static void WriteSharedString(BiffBuffer xml, int sharedStringIndex, int columnIndex, int rowNumber, bool includeReference)
        {
            if (includeReference)
            {
                WriteCellOpenWithRef(xml, columnIndex, rowNumber, " t=\"s\"><v>"u8);
            }
            else
            {
                xml.Write("<c t=\"s\"><v>"u8);
            }
            WriteValue(xml, sharedStringIndex, sizeHint: 16);
            xml.Write("</v></c>"u8);
        }

        internal static void WriteBool(BiffBuffer xml, bool value, int columnIndex, int rowNumber, bool includeReference)
        {
            if (includeReference)
            {
                WriteCellOpenWithRef(xml, columnIndex, rowNumber, " t=\"b\"><v>"u8);
            }
            else
            {
                xml.Write("<c t=\"b\"><v>"u8);
            }
            xml.WriteByte(value ? (byte)'1' : (byte)'0');
            xml.Write("</v></c>"u8);
        }

        internal static void WriteDateTime(BiffBuffer xml, DateTime value, int columnIndex, int rowNumber, bool includeReference)
        {
            if (includeReference)
            {
                WriteCellOpenWithRef(xml, columnIndex, rowNumber, " s=\"1\"><v>"u8);
            }
            else
            {
                xml.Write("<c s=\"1\"><v>"u8);
            }
            WriteValue(xml, DateSerial.ForEpoch(value.ToOADate(), date1904: false), sizeHint: 32);
            xml.Write("</v></c>"u8);
        }

        internal static void WriteNumber<T>(BiffBuffer xml, T value, int columnIndex, int rowNumber, bool includeReference)
            where T : IUtf8SpanFormattable
        {
            if (includeReference)
            {
                WriteCellOpenWithRef(xml, columnIndex, rowNumber, "><v>"u8);
            }
            else
            {
                xml.Write("<c><v>"u8);
            }
            WriteValue(xml, value, sizeHint: 64);
            xml.Write("</v></c>"u8);
        }

        private static bool HasEdgeWhitespace(string value)
        {
            return value.Length != 0 && (char.IsWhiteSpace(value[0]) || char.IsWhiteSpace(value[^1]));
        }

        internal static void WriteNumber(BiffBuffer xml, int value, int columnIndex, int rowNumber, bool includeReference)
        {
            if (includeReference)
            {
                WriteCellOpenWithRef(xml, columnIndex, rowNumber, "><v>"u8);
            }
            else
            {
                xml.Write("<c><v>"u8);
            }
            WriteValue(xml, value, sizeHint: 16);
            xml.Write("</v></c>"u8);
        }

        internal static void WriteNumber(BiffBuffer xml, long value, int columnIndex, int rowNumber, bool includeReference)
        {
            if (includeReference)
            {
                WriteCellOpenWithRef(xml, columnIndex, rowNumber, "><v>"u8);
            }
            else
            {
                xml.Write("<c><v>"u8);
            }
            WriteValue(xml, value, sizeHint: 32);
            xml.Write("</v></c>"u8);
        }

        internal static void WriteNumber(BiffBuffer xml, double value, int columnIndex, int rowNumber, bool includeReference)
        {
            if (includeReference)
            {
                WriteCellOpenWithRef(xml, columnIndex, rowNumber, "><v>"u8);
            }
            else
            {
                xml.Write("<c><v>"u8);
            }
            WriteValue(xml, value, sizeHint: 32);
            xml.Write("</v></c>"u8);
        }

        internal static void WriteNumber(BiffBuffer xml, decimal value, int columnIndex, int rowNumber, bool includeReference)
        {
            if (includeReference)
            {
                WriteCellOpenWithRef(xml, columnIndex, rowNumber, "><v>"u8);
            }
            else
            {
                xml.Write("<c><v>"u8);
            }
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
            ThrowIfNonFinite(value);
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
                ThrowIfNonFinite(Unsafe.As<T, double>(ref value));
            }
            else if (typeof(T) == typeof(float))
            {
                float f = Unsafe.As<T, float>(ref value);
                if (!float.IsFinite(f))
                {
                    throw new ArgumentException($"Cannot write non-finite value '{f}' to a spreadsheet cell.", nameof(value));
                }
            }
            int size = sizeHint;
            int written;
            while (!value.TryFormat(xml.GetSpan(size), out written, default, CultureInfo.InvariantCulture))
            {
                size = checked(size * 2);
            }
            xml.Advance(written);
        }

        // NaN/Infinity have no representation in the numeric <v> element ([ISO/IEC 29500] ST_Xstring
        // doesn't cover it either) — writing them as literal text produces a file Excel rejects on open.
        private static void ThrowIfNonFinite(double value)
        {
            if (!double.IsFinite(value))
            {
                throw new ArgumentException($"Cannot write non-finite value '{value}' to a spreadsheet cell.", nameof(value));
            }
        }

        internal static void WriteEscaped(BiffBuffer xml, ReadOnlySpan<char> value)
        {
            int start = 0;
            int next = value.IndexOfAny(specialChars);
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
                        int following = value[(i + 1)..].IndexOfAny(specialChars);
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
                next = value[start..].IndexOfAny(specialChars);
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
