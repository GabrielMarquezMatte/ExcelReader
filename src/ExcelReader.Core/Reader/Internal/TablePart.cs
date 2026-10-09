using System.Buffers.Text;
using System.Text;
using ExcelReader.Core.Writer.Internal;

namespace ExcelReader.Core.Reader.Internal
{
    internal sealed class TablePart
    {
        private TablePart(int sheetIndex, int id, string name, int firstRow, int firstColumn, int lastRow, int lastColumn,
            int headerRowCount, int totalsRowCount, string[] columns, string? styleName)
        {
            SheetIndex = sheetIndex;
            Id = id;
            Name = name;
            FirstRow = firstRow;
            FirstColumn = firstColumn;
            LastRow = lastRow;
            LastColumn = lastColumn;
            HeaderRowCount = headerRowCount;
            TotalsRowCount = totalsRowCount;
            Columns = columns;
            StyleName = styleName;
        }

        internal int SheetIndex { get; }
        internal int Id { get; }
        internal string Name { get; }
        internal int FirstRow { get; }
        internal int FirstColumn { get; }
        internal int LastRow { get; }
        internal int LastColumn { get; }
        internal int HeaderRowCount { get; }
        internal int TotalsRowCount { get; }
        internal string[] Columns { get; }
        internal string? StyleName { get; }

        internal static TablePart Create(int sheetIndex, int id, string name, int firstRow, int firstColumn, int lastRow, int lastColumn,
            int headerRowCount, int totalsRowCount, string[] columns, string? styleName)
        {
            if (name.Length == 0)
            {
                throw new InvalidDataException("A table has no name.");
            }
            if (firstRow < 0 || firstRow > lastRow || lastRow >= ExcelLimits.MaxRows
                || firstColumn < 0 || firstColumn > lastColumn || lastColumn >= ExcelLimits.MaxColumns)
            {
                throw Malformed(name, "its range lies outside the sheet");
            }
            if (headerRowCount is < 0 or > 1 || totalsRowCount is < 0 or > 1
                || headerRowCount + totalsRowCount > lastRow - firstRow + 1)
            {
                throw Malformed(name, "its header and totals row counts do not fit its range");
            }
            if (columns.Length != lastColumn - firstColumn + 1)
            {
                throw Malformed(name, "its column count does not match its range");
            }
            return new TablePart(sheetIndex, id, name, firstRow, firstColumn, lastRow, lastColumn,
                headerRowCount, totalsRowCount, columns, styleName);
        }

        internal static bool TryParseRange(ReadOnlySpan<byte> text, out int firstRow, out int firstColumn, out int lastRow, out int lastColumn)
        {
            int colon = text.IndexOf((byte)':');
            if (colon < 0)
            {
                bool parsed = TryParseCell(text, out firstRow, out firstColumn);
                lastRow = firstRow;
                lastColumn = firstColumn;
                return parsed;
            }
            lastRow = 0;
            lastColumn = 0;
            return TryParseCell(text[..colon], out firstRow, out firstColumn)
                && TryParseCell(text[(colon + 1)..], out lastRow, out lastColumn);
        }

        internal static string FormatRange(int firstRow, int firstColumn, int lastRow, int lastColumn)
        {
            Span<byte> buffer = stackalloc byte[32];
            int written = WriteCell(buffer, firstRow, firstColumn);
            buffer[written++] = (byte)':';
            written += WriteCell(buffer[written..], lastRow, lastColumn);
            return Encoding.ASCII.GetString(buffer[..written]);
        }

        private static bool TryParseCell(ReadOnlySpan<byte> text, out int row, out int column)
        {
            row = 0;
            column = 0;
            int letters = 0;
            while (letters < text.Length && (uint)(text[letters] - 'A') <= 25)
            {
                column = (column * 26) + (text[letters] - 'A' + 1);
                letters++;
                if (column > ExcelLimits.MaxColumns)
                {
                    return false;
                }
            }
            if (letters == 0 || !Utf8Parser.TryParse(text[letters..], out int rowNumber, out int consumed)
                || consumed != text.Length - letters || rowNumber < 1)
            {
                return false;
            }
            row = rowNumber - 1;
            column--;
            return true;
        }

        private static int WriteCell(Span<byte> destination, int row, int column)
        {
            int written = ColumnName.Write(destination, column);
            Utf8Formatter.TryFormat(row + 1, destination[written..], out int digits);
            return written + digits;
        }

        private static InvalidDataException Malformed(string name, string reason)
        {
            return new InvalidDataException($"Table '{name}' is malformed: {reason}.");
        }
    }
}
