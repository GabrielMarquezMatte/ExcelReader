using System.Buffers;
using System.Buffers.Text;
using System.Globalization;

namespace ExcelReader.Core.Writer.Internal
{
    internal static class CellFormatter
    {
        private static readonly SearchValues<char> specialChars = SearchValues.Create("&<>\"'");

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
                WriteCellOpenWithRef(xml, columnIndex, rowNumber, " t=\"inlineStr\"><is><t>"u8);
            }
            else
            {
                xml.Write("<c t=\"inlineStr\"><is><t>"u8);
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
            WriteValue(xml, value.ToOADate(), sizeHint: 32);
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
            int next = value.IndexOfAny(specialChars);
            while (next >= 0)
            {
                int i = start + next;
                ReadOnlySpan<byte> escape = value[i] switch
                {
                    '&' => "&amp;"u8,
                    '<' => "&lt;"u8,
                    '>' => "&gt;"u8,
                    '"' => "&quot;"u8,
                    _ => "&apos;"u8, // '\''
                };
                if (i > start)
                {
                    xml.WriteUtf8(value[start..i]);
                }
                xml.Write(escape);
                start = i + 1;
                next = value[start..].IndexOfAny(specialChars);
            }
            if (start < value.Length)
            {
                xml.WriteUtf8(value[start..]);
            }
        }
    }
}
