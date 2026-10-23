using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;
using System.Text;
using ExcelReader.Core.Writer;
using ExcelReader.Native.Typed;

namespace ExcelReader.Native.Writer
{
    // See xl_sheet_options in excelreader.h for the authoritative field list.
    [StructLayout(LayoutKind.Sequential)]
    internal unsafe struct NativeSheetOptionsRaw
    {
        public int StructSize;
        public int TableNameLen;
        public byte* TableName;
        public int TableStyleLen;
        public byte* TableStyle;
        public int TableFlags;
    }

    /// <summary>An Excel table to write a sheet as, decoded from <see cref="NativeSheetOptionsRaw"/>.</summary>
    internal sealed record NativeExcelTable(string Name, ExcelTableOptions Options)
    {
        internal const int NoStyle = 1;
        internal const int NoRowStripes = 2;
        internal const int ColumnStripes = 4;
        internal const int FirstColumn = 8;
        internal const int LastColumn = 16;
        private const int KnownFlags = NoStyle | NoRowStripes | ColumnStripes | FirstColumn | LastColumn;

        internal static unsafe bool TryRead(NativeSheetOptionsRaw* options, out NativeExcelTable? table, [NotNullWhen(false)] out string? error)
        {
            table = null;
            error = null;
            if (options is null)
            {
                return true;
            }
            if (options->StructSize != sizeof(NativeSheetOptionsRaw))
            {
                error = $"xl_sheet_options.struct_size is {options->StructSize}, but this library expects {sizeof(NativeSheetOptionsRaw)}.";
                return false;
            }
            int flags = options->TableFlags;
            if (options->TableName is null)
            {
                if (options->TableStyle is not null || flags != 0)
                {
                    error = "xl_sheet_options sets a table style or flags without a table_name.";
                    return false;
                }
                return true;
            }
            if ((flags & ~KnownFlags) != 0)
            {
                error = $"xl_sheet_options.table_flags has unknown bits; got {flags}.";
                return false;
            }
            if (!TypedApi.IsValidNameLength(options->TableNameLen))
            {
                error = $"xl_sheet_options.table_name_len is out of range; got {options->TableNameLen}.";
                return false;
            }
            string? style = ExcelTableOptions.Default.StyleName;
            if (options->TableStyle is not null)
            {
                if ((flags & NoStyle) != 0 || !TypedApi.IsValidNameLength(options->TableStyleLen))
                {
                    error = "xl_sheet_options.table_style needs a valid table_style_len and cannot be combined with XL_TABLE_NO_STYLE.";
                    return false;
                }
                style = Encoding.UTF8.GetString(options->TableStyle, options->TableStyleLen);
            }
            table = new NativeExcelTable(Encoding.UTF8.GetString(options->TableName, options->TableNameLen), new ExcelTableOptions
            {
                StyleName = (flags & NoStyle) != 0 ? null : style,
                ShowRowStripes = (flags & NoRowStripes) == 0,
                ShowColumnStripes = (flags & ColumnStripes) != 0,
                ShowFirstColumn = (flags & FirstColumn) != 0,
                ShowLastColumn = (flags & LastColumn) != 0,
            });
            return true;
        }
    }
}
