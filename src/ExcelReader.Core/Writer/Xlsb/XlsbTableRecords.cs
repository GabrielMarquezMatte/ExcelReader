using ExcelReader.Core.Reader.Xlsb;
using ExcelReader.Core.Writer.Internal;

namespace ExcelReader.Core.Writer.Xlsb
{
    internal static class XlsbTableRecords
    {
        private const uint Unset = 0xFFFFFFFF;
        private const uint RangeListType = 0;
        private const uint HeaderRowCount = 1;
        private const uint TotalsRowCount = 0;
        private const uint ListFlags = 0;
        private const uint ConnectionId = 0;
        private const uint NoTotalsFunction = 0;
        private const uint NoQuerySource = 0;
        private const int ListDxfCount = 6;
        private const int ColumnDxfCount = 3;
        private const int ListTrailingStrings = 3;
        private const int ColumnTrailingStrings = 4;
        private const int FirstColumnFlag = 1;
        private const int LastColumnFlag = 2;
        private const int RowStripesFlag = 4;
        private const int ColumnStripesFlag = 8;

        internal static void WriteTable(BiffBuffer dest, BiffBuffer payload, WrittenTable table)
        {
            WriteBeginList(dest, payload, table);
            payload.Reset();
            WriteRange(payload, table);
            Biff12RecordWriter.WriteRecord(dest, Brt.BeginAFilter, payload.Span);
            Biff12RecordWriter.WriteRecord(dest, Brt.EndAFilter);
            WriteColumns(dest, payload, table.Columns);
            WriteStyle(dest, payload, table.Options);
            Biff12RecordWriter.WriteRecord(dest, Brt.EndList);
        }

        internal static int StyleFlags(ExcelTableOptions options)
        {
            int flags = 0;
            if (options.ShowFirstColumn)
            {
                flags |= FirstColumnFlag;
            }
            if (options.ShowLastColumn)
            {
                flags |= LastColumnFlag;
            }
            if (options.ShowRowStripes)
            {
                flags |= RowStripesFlag;
            }
            if (options.ShowColumnStripes)
            {
                flags |= ColumnStripesFlag;
            }
            return flags;
        }

        private static void WriteBeginList(BiffBuffer dest, BiffBuffer payload, WrittenTable table)
        {
            payload.Reset();
            WriteRange(payload, table);
            payload.WriteU32(RangeListType);
            payload.WriteU32((uint)table.Id);
            payload.WriteU32(HeaderRowCount);
            payload.WriteU32(TotalsRowCount);
            payload.WriteU32(ListFlags);
            WriteUnset(payload, ListDxfCount);
            payload.WriteU32(ConnectionId);
            Biff12RecordWriter.WriteWideString(payload, table.Name);
            Biff12RecordWriter.WriteWideString(payload, table.Name);
            Biff12RecordWriter.WriteWideString(payload, ReadOnlySpan<char>.Empty);
            WriteUnset(payload, ListTrailingStrings);
            Biff12RecordWriter.WriteRecord(dest, Brt.BeginList, payload.Span);
        }

        private static void WriteRange(BiffBuffer payload, WrittenTable table)
        {
            payload.WriteU32((uint)table.HeaderRow);
            payload.WriteU32((uint)table.LastRow);
            payload.WriteU32((uint)table.FirstColumn);
            payload.WriteU32((uint)table.LastColumn);
        }

        private static void WriteColumns(BiffBuffer dest, BiffBuffer payload, string[] columns)
        {
            payload.Reset();
            payload.WriteU32((uint)columns.Length);
            Biff12RecordWriter.WriteRecord(dest, Brt.BeginListCols, payload.Span);
            for (int i = 0; i < columns.Length; i++)
            {
                payload.Reset();
                payload.WriteU32((uint)(i + 1));
                payload.WriteU32(NoTotalsFunction);
                WriteUnset(payload, ColumnDxfCount);
                payload.WriteU32(NoQuerySource);
                payload.WriteU32(Unset);
                Biff12RecordWriter.WriteWideString(payload, columns[i]);
                WriteUnset(payload, ColumnTrailingStrings);
                Biff12RecordWriter.WriteRecord(dest, Brt.BeginListCol, payload.Span);
                Biff12RecordWriter.WriteRecord(dest, Brt.EndListCol);
            }
            Biff12RecordWriter.WriteRecord(dest, Brt.EndListCols);
        }

        private static void WriteStyle(BiffBuffer dest, BiffBuffer payload, ExcelTableOptions options)
        {
            payload.Reset();
            payload.WriteU16(StyleFlags(options));
            if (options.StyleName is null)
            {
                payload.WriteU32(Unset);
            }
            else
            {
                Biff12RecordWriter.WriteWideString(payload, options.StyleName);
            }
            Biff12RecordWriter.WriteRecord(dest, Brt.ListTableStyleClient, payload.Span);
        }

        private static void WriteUnset(BiffBuffer payload, int count)
        {
            for (int i = 0; i < count; i++)
            {
                payload.WriteU32(Unset);
            }
        }
    }
}
