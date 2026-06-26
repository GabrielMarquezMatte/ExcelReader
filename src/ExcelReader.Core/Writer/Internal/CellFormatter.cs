using System.Buffers;
using System.Buffers.Text;
using System.Globalization;

namespace ExcelReader.Core.Writer.Internal
{
    internal static class CellFormatter
    {
        // Writes the cell reference (e.g. "B7") directly to the writer.
        // Max XLSX cell is XFD1048576 -> 3 column letters + 7 row digits.
        private static void WriteRef(BiffBuffer xml, int columnIndex, int rowNumber)
        {
            Span<byte> buf = stackalloc byte[10];
            int len = ColumnName.Write(buf, columnIndex);
            Utf8Formatter.TryFormat(rowNumber, buf[len..], out int rowLen);
            xml.Write(buf[..(len + rowLen)]);
        }

        private static void WriteCellOpen(BiffBuffer xml, int columnIndex, int rowNumber, bool includeReference)
        {
            xml.Write("<c"u8);
            if (!includeReference)
            {
                return;
            }
            xml.Write(" r=\""u8);
            WriteRef(xml, columnIndex, rowNumber);
            xml.WriteByte((byte)'"');
        }

        internal static void WriteEmpty(BiffBuffer xml, int columnIndex, int rowNumber, bool includeReference)
        {
            WriteCellOpen(xml, columnIndex, rowNumber, includeReference);
            xml.Write("/>"u8);
        }

        internal static void WriteString(BiffBuffer xml, string value, int columnIndex, int rowNumber, bool includeReference)
        {
            WriteCellOpen(xml, columnIndex, rowNumber, includeReference);
            xml.Write(" t=\"inlineStr\"><is><t>"u8);
            WriteEscaped(xml, value);
            xml.Write("</t></is></c>"u8);
        }

        internal static void WriteBool(BiffBuffer xml, bool value, int columnIndex, int rowNumber, bool includeReference)
        {
            WriteCellOpen(xml, columnIndex, rowNumber, includeReference);
            xml.Write(" t=\"b\"><v>"u8);
            xml.WriteByte(value ? (byte)'1' : (byte)'0');
            xml.Write("</v></c>"u8);
        }

        internal static void WriteDateTime(BiffBuffer xml, DateTime value, int columnIndex, int rowNumber, bool includeReference)
        {
            WriteCellOpen(xml, columnIndex, rowNumber, includeReference);
            xml.Write(" s=\"1\"><v>"u8);
            double oaDate = value.ToOADate();
            WriteDoubleValue(xml, oaDate);
            xml.Write("</v></c>"u8);
        }

        internal static void WriteNumber<T>(BiffBuffer xml, T value, int columnIndex, int rowNumber, bool includeReference)
            where T : ISpanFormattable
        {
            WriteCellOpen(xml, columnIndex, rowNumber, includeReference);
            xml.Write("><v>"u8);
            Span<char> buf = stackalloc char[64];
            value.TryFormat(buf, out int written, default, CultureInfo.InvariantCulture);
            xml.WriteUtf8(buf[..written]);
            xml.Write("</v></c>"u8);
        }

        internal static void WriteNumber(BiffBuffer xml, int value, int columnIndex, int rowNumber, bool includeReference)
        {
            WriteCellOpen(xml, columnIndex, rowNumber, includeReference);
            xml.Write("><v>"u8);
            WriteIntValue(xml, value);
            xml.Write("</v></c>"u8);
        }

        internal static void WriteNumber(BiffBuffer xml, long value, int columnIndex, int rowNumber, bool includeReference)
        {
            WriteCellOpen(xml, columnIndex, rowNumber, includeReference);
            xml.Write("><v>"u8);
            WriteLongValue(xml, value);
            xml.Write("</v></c>"u8);
        }

        internal static void WriteNumber(BiffBuffer xml, double value, int columnIndex, int rowNumber, bool includeReference)
        {
            WriteCellOpen(xml, columnIndex, rowNumber, includeReference);
            xml.Write("><v>"u8);
            WriteDoubleValue(xml, value);
            xml.Write("</v></c>"u8);
        }

        internal static void WriteNumber(BiffBuffer xml, decimal value, int columnIndex, int rowNumber, bool includeReference)
        {
            WriteCellOpen(xml, columnIndex, rowNumber, includeReference);
            xml.Write("><v>"u8);
            WriteDecimalValue(xml, value);
            xml.Write("</v></c>"u8);
        }

        private static void WriteIntValue(BiffBuffer xml, int value)
        {
            Span<byte> buf = stackalloc byte[16];
            Utf8Formatter.TryFormat(value, buf, out int written);
            xml.Write(buf[..written]);
        }

        private static void WriteLongValue(BiffBuffer xml, long value)
        {
            Span<byte> buf = stackalloc byte[32];
            Utf8Formatter.TryFormat(value, buf, out int written);
            xml.Write(buf[..written]);
        }

        private static void WriteDoubleValue(BiffBuffer xml, double value)
        {
            Span<byte> buf = stackalloc byte[32];
            Utf8Formatter.TryFormat(value, buf, out int written, new StandardFormat('G', 17));
            xml.Write(buf[..written]);
        }

        private static void WriteDecimalValue(BiffBuffer xml, decimal value)
        {
            Span<byte> buf = stackalloc byte[64];
            Utf8Formatter.TryFormat(value, buf, out int written);
            xml.Write(buf[..written]);
        }

        private static void WriteEscaped(BiffBuffer xml, ReadOnlySpan<char> value)
        {
            int start = 0;
            for (int i = 0; i < value.Length; i++)
            {
                ReadOnlySpan<byte> escape = value[i] switch
                {
                    '&' => "&amp;"u8,
                    '<' => "&lt;"u8,
                    '>' => "&gt;"u8,
                    '"' => "&quot;"u8,
                    '\'' => "&apos;"u8,
                    _ => default,
                };
                if (escape.IsEmpty)
                {
                    continue;
                }
                if (i > start)
                {
                    xml.WriteUtf8(value[start..i]);
                }
                xml.Write(escape);
                start = i + 1;
            }
            if (start < value.Length)
            {
                xml.WriteUtf8(value[start..]);
            }
        }
    }
}
