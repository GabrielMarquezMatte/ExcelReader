using System.Runtime.InteropServices;

namespace ExcelReader.NetFramework.Interop
{
    internal static unsafe class NativeMethods
    {
        // Resolves only after NativeLoader.EnsureLoaded has loaded the DLL by absolute path.
        private const string Lib = "ExcelReader.Native.dll";

        [DllImport(Lib, CallingConvention = CallingConvention.Cdecl, ExactSpelling = true)]
        internal static extern int xl_abi_version();

        [DllImport(Lib, CallingConvention = CallingConvention.Cdecl, ExactSpelling = true)]
        internal static extern byte* xl_last_error_ptr(int* outLen);

        [DllImport(Lib, CallingConvention = CallingConvention.Cdecl, ExactSpelling = true)]
        internal static extern int xl_open_file(byte* path, int pathLen, int format, XlOpenOptions* options, IntPtr* outHandle);

        [DllImport(Lib, CallingConvention = CallingConvention.Cdecl, ExactSpelling = true)]
        internal static extern int xl_open_memory(byte* data, int dataLen, int format, XlOpenOptions* options, IntPtr* outHandle);

        [DllImport(Lib, CallingConvention = CallingConvention.Cdecl, ExactSpelling = true)]
        internal static extern int xl_open_stream(XlStream* stream, int format, XlOpenOptions* options, IntPtr* outHandle);

        [DllImport(Lib, CallingConvention = CallingConvention.Cdecl, ExactSpelling = true)]
        internal static extern void xl_set_source_error(byte* message, int len);

        [DllImport(Lib, CallingConvention = CallingConvention.Cdecl, ExactSpelling = true)]
        internal static extern int xl_close(IntPtr handle);

        [DllImport(Lib, CallingConvention = CallingConvention.Cdecl, ExactSpelling = true)]
        internal static extern int xl_sheet_count(IntPtr handle, int* outCount);

        [DllImport(Lib, CallingConvention = CallingConvention.Cdecl, ExactSpelling = true)]
        internal static extern int xl_sheet_name_at(IntPtr handle, int index, byte* buffer, int capacity, int* outLen);

        [DllImport(Lib, CallingConvention = CallingConvention.Cdecl, ExactSpelling = true)]
        internal static extern int xl_sheet_visibility_at(IntPtr handle, int index, int* outVisibility);

        [DllImport(Lib, CallingConvention = CallingConvention.Cdecl, ExactSpelling = true)]
        internal static extern int xl_sheet_index(IntPtr handle, byte* name, int nameLen, int* outIndex);

        [DllImport(Lib, CallingConvention = CallingConvention.Cdecl, ExactSpelling = true)]
        internal static extern int xl_is_date1904(IntPtr handle, int* outFlag);

        [DllImport(Lib, CallingConvention = CallingConvention.Cdecl, ExactSpelling = true)]
        internal static extern int xl_table_count(IntPtr handle, int* outCount);

        [DllImport(Lib, CallingConvention = CallingConvention.Cdecl, ExactSpelling = true)]
        internal static extern int xl_table_info_at(IntPtr handle, int index, XlTableInfo* outInfo);

        [DllImport(Lib, CallingConvention = CallingConvention.Cdecl, ExactSpelling = true)]
        internal static extern int xl_table_name_at(IntPtr handle, int index, byte* buffer, int capacity, int* outLen);

        [DllImport(Lib, CallingConvention = CallingConvention.Cdecl, ExactSpelling = true)]
        internal static extern int xl_table_index(IntPtr handle, byte* name, int nameLen, int* outIndex);

        [DllImport(Lib, CallingConvention = CallingConvention.Cdecl, ExactSpelling = true)]
        internal static extern int xl_rows_open(IntPtr handle, int sheet, IntPtr* outRows);

        [DllImport(Lib, CallingConvention = CallingConvention.Cdecl, ExactSpelling = true)]
        internal static extern int xl_rows_next_view(IntPtr rows, XlRow* outRow);

        [DllImport(Lib, CallingConvention = CallingConvention.Cdecl, ExactSpelling = true)]
        internal static extern int xl_rows_close(IntPtr rows);

        [DllImport(Lib, CallingConvention = CallingConvention.Cdecl, ExactSpelling = true)]
        internal static extern int xl_rows_foreach(IntPtr handle, int sheet, IntPtr callback, IntPtr userData);

        [DllImport(Lib, CallingConvention = CallingConvention.Cdecl, ExactSpelling = true)]
        internal static extern int xl_aggregate(IntPtr handle, int sheet, XlAggregation* aggregation, int headerRow, int degreeOfParallelism, IntPtr* outState);

        [DllImport(Lib, CallingConvention = CallingConvention.Cdecl, ExactSpelling = true)]
        internal static extern int xl_parse_typed(IntPtr handle, int sheet, XlColumnSpec* specs, int specCount, int headerRow, int degreeOfParallelism, XlTable* outTable);

        [DllImport(Lib, CallingConvention = CallingConvention.Cdecl, ExactSpelling = true)]
        internal static extern void xl_free_table(XlTable* table);

        [DllImport(Lib, CallingConvention = CallingConvention.Cdecl, ExactSpelling = true)]
        internal static extern int xl_typed_reader_open(IntPtr handle, int sheet, XlColumnSpec* specs, int specCount, int headerRow, long maxRows, IntPtr* outReader);

        [DllImport(Lib, CallingConvention = CallingConvention.Cdecl, ExactSpelling = true)]
        internal static extern int xl_typed_reader_next(IntPtr reader, XlTable* outTable);

        [DllImport(Lib, CallingConvention = CallingConvention.Cdecl, ExactSpelling = true)]
        internal static extern void xl_typed_reader_close(IntPtr reader);

        [DllImport(Lib, CallingConvention = CallingConvention.Cdecl, ExactSpelling = true)]
        internal static extern int xl_infer_schema(IntPtr handle, int sheet, int headerRow, int sampleSize, int flags, XlInferredSchema* outSchema);

        [DllImport(Lib, CallingConvention = CallingConvention.Cdecl, ExactSpelling = true)]
        internal static extern void xl_free_schema(XlInferredSchema* schema);

        [DllImport(Lib, CallingConvention = CallingConvention.Cdecl, ExactSpelling = true)]
        internal static extern int xl_open_write_handle(byte* path, int pathLen, int format, XlWriteOptions* options, IntPtr* outHandle);

        [DllImport(Lib, CallingConvention = CallingConvention.Cdecl, ExactSpelling = true)]
        internal static extern int xl_open_write_handle_to_memory(int format, XlWriteOptions* options, IntPtr* outHandle);

        [DllImport(Lib, CallingConvention = CallingConvention.Cdecl, ExactSpelling = true)]
        internal static extern int xl_start_sheet(IntPtr handle, byte* name, int nameLen);

        [DllImport(Lib, CallingConvention = CallingConvention.Cdecl, ExactSpelling = true)]
        internal static extern int xl_start_row(IntPtr handle);

        [DllImport(Lib, CallingConvention = CallingConvention.Cdecl, ExactSpelling = true)]
        internal static extern int xl_add_style(IntPtr handle, XlCellStyle* style, int* outStyleId);

        [DllImport(Lib, CallingConvention = CallingConvention.Cdecl, ExactSpelling = true)]
        internal static extern int xl_set_column_style(IntPtr handle, int column, int styleId);

        [DllImport(Lib, CallingConvention = CallingConvention.Cdecl, ExactSpelling = true)]
        internal static extern int xl_set_column_width(IntPtr handle, int column, double width);

        [DllImport(Lib, CallingConvention = CallingConvention.Cdecl, ExactSpelling = true)]
        internal static extern int xl_start_row_styled(IntPtr handle, int styleId, double height);

        [DllImport(Lib, CallingConvention = CallingConvention.Cdecl, ExactSpelling = true)]
        internal static extern int xl_style_next_cell(IntPtr handle, int styleId);

        [DllImport(Lib, CallingConvention = CallingConvention.Cdecl, ExactSpelling = true)]
        internal static extern int xl_write_string(IntPtr handle, byte* value, int valueLen);

        [DllImport(Lib, CallingConvention = CallingConvention.Cdecl, ExactSpelling = true)]
        internal static extern int xl_write_int64(IntPtr handle, long value);

        [DllImport(Lib, CallingConvention = CallingConvention.Cdecl, ExactSpelling = true)]
        internal static extern int xl_write_float64(IntPtr handle, double value);

        [DllImport(Lib, CallingConvention = CallingConvention.Cdecl, ExactSpelling = true)]
        internal static extern int xl_write_bool(IntPtr handle, int value);

        [DllImport(Lib, CallingConvention = CallingConvention.Cdecl, ExactSpelling = true)]
        internal static extern int xl_write_date(IntPtr handle, int daysSinceEpoch);

        [DllImport(Lib, CallingConvention = CallingConvention.Cdecl, ExactSpelling = true)]
        internal static extern int xl_write_time(IntPtr handle, long microsecondsSinceMidnight);

        [DllImport(Lib, CallingConvention = CallingConvention.Cdecl, ExactSpelling = true)]
        internal static extern int xl_write_timestamp(IntPtr handle, long microsecondsSinceEpoch);

        [DllImport(Lib, CallingConvention = CallingConvention.Cdecl, ExactSpelling = true)]
        internal static extern int xl_write_null(IntPtr handle, int type);

        [DllImport(Lib, CallingConvention = CallingConvention.Cdecl, ExactSpelling = true)]
        internal static extern int xl_end_row(IntPtr handle);

        [DllImport(Lib, CallingConvention = CallingConvention.Cdecl, ExactSpelling = true)]
        internal static extern int xl_end_sheet(IntPtr handle);

        [DllImport(Lib, CallingConvention = CallingConvention.Cdecl, ExactSpelling = true)]
        internal static extern int xl_write_sheet(IntPtr handle, byte* name, int nameLen, XlColumnSpec* specs, XlTable* table, XlSheetOptions* options);

        [DllImport(Lib, CallingConvention = CallingConvention.Cdecl, ExactSpelling = true)]
        internal static extern int xl_close_write_handle(IntPtr handle);

        [DllImport(Lib, CallingConvention = CallingConvention.Cdecl, ExactSpelling = true)]
        internal static extern int xl_write_handle_bytes(IntPtr handle, XlBuffer* outBuffer);

        [DllImport(Lib, CallingConvention = CallingConvention.Cdecl, ExactSpelling = true)]
        internal static extern int xl_abort_write_handle(IntPtr handle);

        [DllImport(Lib, CallingConvention = CallingConvention.Cdecl, ExactSpelling = true)]
        internal static extern void xl_free_buffer(XlBuffer* buffer);

        [DllImport(Lib, CallingConvention = CallingConvention.Cdecl, ExactSpelling = true)]
        internal static extern int xl_encrypt_package_to_memory(byte* package, int packageLen, byte* password, int passwordLen, XlBuffer* outBuffer);
    }
}
