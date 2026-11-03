using System.Runtime.InteropServices;

namespace ExcelReader.NetFramework.Interop
{
    internal static unsafe class NativeMethods
    {
        private const string Lib = "ExcelReader.Native.dll";

        static NativeMethods() => NativeLoader.EnsureLoaded();

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
    }
}
