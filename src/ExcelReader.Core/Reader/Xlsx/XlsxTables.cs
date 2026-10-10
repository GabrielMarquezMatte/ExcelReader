using System.Buffers.Text;
using ExcelReader.Core.Reader.Internal;

namespace ExcelReader.Core.Reader.Xlsx
{
    internal static class XlsxTables
    {
        internal static TablePart ParseTable(int sheetIndex, ReadOnlySpan<byte> xml)
        {
            ReadOnlySpan<byte> prefix = XlsxXml.DetectElementPrefix(xml);
            ReadOnlySpan<byte> table = FirstTag(xml, XlsxXml.Token("<"u8, prefix, "table"u8));
            if (table.IsEmpty)
            {
                throw new InvalidDataException("A table part has no <table> element.");
            }
            string name = XlsxXml.DecodeToString(XlsxXml.Attr(table, " displayName="u8));
            if (!TablePart.TryParseRange(XlsxXml.Attr(table, " ref="u8), out int firstRow, out int firstColumn, out int lastRow, out int lastColumn))
            {
                throw new InvalidDataException($"Table '{name}' has an unreadable ref.");
            }
            int id = XlsxXml.ParseIntOr(XlsxXml.Attr(table, " id="u8), 0);
            int headerRowCount = ParseRowCount(table, " headerRowCount="u8, 1, name);
            int totalsRowCount = ParseRowCount(table, " totalsRowCount="u8, 0, name);
            return TablePart.Create(sheetIndex, id, name, firstRow, firstColumn, lastRow, lastColumn,
                headerRowCount, totalsRowCount, ColumnNames(xml, prefix), StyleName(xml, prefix));
        }

        private static int ParseRowCount(ReadOnlySpan<byte> table, ReadOnlySpan<byte> attribute, int fallback, string name)
        {
            ReadOnlySpan<byte> value = XlsxXml.Attr(table, attribute);
            if (value.IsEmpty)
            {
                return fallback;
            }
            if (!Utf8Parser.TryParse(value, out int count, out int consumed) || consumed != value.Length)
            {
                throw new InvalidDataException($"Table '{name}' has an unreadable row count.");
            }
            return count;
        }

        private static string[] ColumnNames(ReadOnlySpan<byte> xml, ReadOnlySpan<byte> prefix)
        {
            List<string> names = [];
            foreach (ReadOnlySpan<byte> column in new TagSpanEnumerable(xml, XlsxXml.Token("<"u8, prefix, "tableColumn"u8)))
            {
                names.Add(XlsxXml.DecodeToString(XlsxXml.Attr(column, " name="u8)));
            }
            return [.. names];
        }

        private static string? StyleName(ReadOnlySpan<byte> xml, ReadOnlySpan<byte> prefix)
        {
            ReadOnlySpan<byte> style = XlsxXml.Attr(FirstTag(xml, XlsxXml.Token("<"u8, prefix, "tableStyleInfo"u8)), " name="u8);
            if (style.IsEmpty)
            {
                return null;
            }
            return XlsxXml.DecodeToString(style);
        }

        private static ReadOnlySpan<byte> FirstTag(ReadOnlySpan<byte> xml, ReadOnlySpan<byte> token)
        {
            foreach (ReadOnlySpan<byte> tag in new TagSpanEnumerable(xml, token))
            {
                return tag;
            }
            return default;
        }
    }
}
