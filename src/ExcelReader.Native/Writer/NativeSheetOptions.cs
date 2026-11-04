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
        public NativeCellStyleRaw** ColumnStyles;
        public double* ColumnWidths;
        public NativeCellStyleRaw* HeaderStyle;

        internal const int V6Size = 40;

        internal static bool TryRead(NativeSheetOptionsRaw* options, out NativeSheetOptionsRaw raw, [NotNullWhen(false)] out string? error)
        {
            raw = default;
            error = null;
            if (options is null)
            {
                return true;
            }
            int size = options->StructSize;
            if (size != sizeof(NativeSheetOptionsRaw) && size != V6Size)
            {
                error = $"xl_sheet_options.struct_size is {size}, but this library expects {sizeof(NativeSheetOptionsRaw)} (or {V6Size} for the ABI 6 layout).";
                return false;
            }
            NativeSheetOptionsRaw copy = default;
            Buffer.MemoryCopy(options, &copy, sizeof(NativeSheetOptionsRaw), size);
            raw = copy;
            return true;
        }
    }

    /// <summary>The column styles, column widths and header style of a typed sheet write; null entries keep the default.</summary>
    internal sealed record NativeSheetStyles(CellStyle?[]? ColumnStyles, double[]? ColumnWidths, CellStyle? HeaderStyle)
    {
        private const double MaxColumnWidth = 255;

        internal static unsafe bool TryRead(in NativeSheetOptionsRaw raw, int columnCount, out NativeSheetStyles? styles, [NotNullWhen(false)] out string? error)
        {
            styles = null;
            error = null;
            if (!NativeCellStyle.TryDecode(raw.HeaderStyle, out CellStyle? header, out string? headerError))
            {
                error = $"xl_sheet_options.header_style: {headerError}";
                return false;
            }
            if (!TryReadColumnStyles(raw.ColumnStyles, columnCount, out CellStyle?[]? columnStyles, out error)
                || !TryReadColumnWidths(raw.ColumnWidths, columnCount, out double[]? widths, out error))
            {
                return false;
            }
            if (columnStyles is not null || widths is not null || header is not null)
            {
                styles = new NativeSheetStyles(columnStyles, widths, header);
            }
            return true;
        }

        private static unsafe bool TryReadColumnStyles(NativeCellStyleRaw** raw, int columnCount, out CellStyle?[]? styles, [NotNullWhen(false)] out string? error)
        {
            styles = null;
            error = null;
            if (raw is null)
            {
                return true;
            }
            CellStyle?[] decoded = new CellStyle?[columnCount];
            for (int index = 0; index < columnCount; index++)
            {
                if (!NativeCellStyle.TryDecode(raw[index], out decoded[index], out string? styleError))
                {
                    error = $"xl_sheet_options.column_styles[{index}]: {styleError}";
                    return false;
                }
            }
            styles = decoded;
            return true;
        }

        private static unsafe bool TryReadColumnWidths(double* raw, int columnCount, out double[]? widths, [NotNullWhen(false)] out string? error)
        {
            widths = null;
            error = null;
            if (raw is null)
            {
                return true;
            }
            double[] copy = new ReadOnlySpan<double>(raw, columnCount).ToArray();
            for (int index = 0; index < columnCount; index++)
            {
                if (!(copy[index] >= 0 && copy[index] <= MaxColumnWidth))
                {
                    error = $"xl_sheet_options.column_widths[{index}] must be 0-{MaxColumnWidth} characters; got {copy[index]}.";
                    return false;
                }
            }
            widths = copy;
            return true;
        }
    }

    /// <summary><see cref="NativeSheetStyles"/> registered with a workbook; style ids below 0 and widths of 0 keep the default.</summary>
    internal readonly record struct SheetStyleIds(int[] TemporalStyleIds, int[]? ColumnStyleIds, double[]? ColumnWidths, int HeaderStyleId);

    /// <summary>An Excel table to write a sheet as, decoded from <see cref="NativeSheetOptionsRaw"/>.</summary>
    internal sealed record NativeExcelTable(string Name, ExcelTableOptions Options)
    {
        internal const int NoStyle = 1;
        internal const int NoRowStripes = 2;
        internal const int ColumnStripes = 4;
        internal const int FirstColumn = 8;
        internal const int LastColumn = 16;
        private const int KnownFlags = NoStyle | NoRowStripes | ColumnStripes | FirstColumn | LastColumn;

        internal static unsafe bool TryRead(in NativeSheetOptionsRaw options, out NativeExcelTable? table, [NotNullWhen(false)] out string? error)
        {
            table = null;
            error = null;
            int flags = options.TableFlags;
            if (options.TableName is null)
            {
                if (options.TableStyle is not null || flags != 0)
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
            if (!TypedApi.IsValidNameLength(options.TableNameLen))
            {
                error = $"xl_sheet_options.table_name_len is out of range; got {options.TableNameLen}.";
                return false;
            }
            string? style = ExcelTableOptions.Default.StyleName;
            if (options.TableStyle is not null)
            {
                if ((flags & NoStyle) != 0 || !TypedApi.IsValidNameLength(options.TableStyleLen))
                {
                    error = "xl_sheet_options.table_style needs a valid table_style_len and cannot be combined with XL_TABLE_NO_STYLE.";
                    return false;
                }
                style = Encoding.UTF8.GetString(options.TableStyle, options.TableStyleLen);
            }
            table = new NativeExcelTable(Encoding.UTF8.GetString(options.TableName, options.TableNameLen), new ExcelTableOptions
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
