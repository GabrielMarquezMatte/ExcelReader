using System.Runtime.InteropServices;

namespace ExcelReader.NetFramework.Interop
{
    internal static class Xl
    {
        internal const int XL_FORMAT_AUTO = 0;
        internal const int XL_FORMAT_XLS = 1;
        internal const int XL_FORMAT_XLSX = 2;
        internal const int XL_FORMAT_XLSB = 3;
        internal const int XL_FORMAT_CSV = 4;

        internal const int XL_T_STRING = 0;
        internal const int XL_T_I64 = 1;
        internal const int XL_T_F64 = 2;
        internal const int XL_T_BOOL = 3;
        internal const int XL_T_DATE = 4;
        internal const int XL_T_TIME = 5;
        internal const int XL_T_TIMESTAMP = 6;
        internal const int XL_INFER_PARSE_TEXT = 1;

        internal const int XL_CELL_EMPTY = 0;
        internal const int XL_CELL_STRING = 1;
        internal const int XL_CELL_NUMBER = 2;
        internal const int XL_CELL_DATE = 3;
        internal const int XL_CELL_BOOL = 4;
        internal const int XL_CELL_FORMULA = 5;
        internal const int XL_CELL_ERROR = 6;

        internal const int XL_SHEET_VISIBLE = 0;
        internal const int XL_SHEET_HIDDEN = 1;
        internal const int XL_SHEET_VERY_HIDDEN = 2;

        internal const int XL_OPT_DEFAULT = 0;
        internal const int XL_OPT_FALSE = 1;
        internal const int XL_OPT_TRUE = 2;

        internal static int TableSheet(int table)
        {
            return -2 - table;
        }
    }

    [StructLayout(LayoutKind.Sequential)]
    internal unsafe struct XlRowCell
    {
        internal int Column;
        internal int Type;
        internal int ValueLen;
        internal byte* Value;
    }

    [StructLayout(LayoutKind.Sequential)]
    internal unsafe struct XlRow
    {
        internal int CellCount;
        internal XlRowCell* Cells;
    }

    [StructLayout(LayoutKind.Sequential)]
    internal unsafe struct XlOpenOptions
    {
        internal int StructSize;
        internal int CsvSniffDialect;
        internal int CsvDelimiter;
        internal int CsvQuote;
        internal int CsvDetectBom;
        internal int CsvMaxCellBytes;
        internal int CsvInternStrings;
        internal long MaxTotalDecompressedBytes;
        internal int MaxCellBytes;
        internal long MaxSharedStringBytes;
        internal int MaxZipEntries;
        internal int PrefetchDecompression;
        internal int InternStrings;
        internal byte* Password;
        internal int PasswordLen;
        internal long SourceCacheBytes;
        internal long MaxBufferedBytes;
        internal int SourceBlockSize;
        internal int VerifyZipChecksums;
    }

    [StructLayout(LayoutKind.Sequential)]
    internal struct XlStream
    {
        internal int StructSize;
        internal IntPtr UserData;
        internal IntPtr Read;
        internal IntPtr Release;
    }

    [StructLayout(LayoutKind.Sequential)]
    internal struct XlTableInfo
    {
        internal int Sheet;
        internal int FirstRow;
        internal int LastRow;
        internal int FirstColumn;
        internal int LastColumn;
        internal int HeaderRowCount;
        internal int TotalsRowCount;
    }

    [StructLayout(LayoutKind.Sequential)]
    internal unsafe struct XlColumnSpec
    {
        internal byte** Names;
        internal int* NameLens;
        internal int NameCount;
        internal int Index;
        internal int Type;
        internal int Nullable;
    }

    [StructLayout(LayoutKind.Sequential)]
    internal unsafe struct XlColumn
    {
        internal int Type;
        internal long Length;
        internal void* Values;
        internal byte* Validity;
        internal byte* Data;
        internal long DataLen;

        internal readonly bool IsValid(int row)
        {
            return Validity is null || (Validity[row >> 3] & (1 << (row & 7))) != 0;
        }
    }

    [StructLayout(LayoutKind.Sequential)]
    internal unsafe struct XlTable
    {
        internal int ColumnCount;
        internal long RowCount;
        internal XlColumn* Columns;
    }

    [StructLayout(LayoutKind.Sequential)]
    internal unsafe struct XlInferredSchema
    {
        internal XlColumnSpec* Columns;
        internal int ColumnCount;
    }
}
