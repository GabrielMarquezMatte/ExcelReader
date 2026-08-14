namespace ExcelReader.Native
{
    /// <summary>Status codes returned by every exported C function. Mirrors include/excelreader.h.</summary>
    internal static class NativeStatus
    {
        internal const int Ok = 0;
        internal const int Eof = -1;
        internal const int BufferTooSmall = -2;
        internal const int InvalidHandle = -3;
        internal const int InvalidArgument = -4;
        internal const int Error = -5;

        /// <summary>ABI revision returned by <c>xl_abi_version</c>. Mirrors XL_ABI_VERSION in include/excelreader.h.
        /// Bump on any change to a struct layout, a status code, or the meaning of an existing function; adding a
        /// new function does not bump it.</summary>
        internal const int AbiVersion = 1;
    }

    /// <summary>
    /// Format selectors accepted by the open functions. Values 1-3 deliberately match
    /// <see cref="Core.Enums.ExcelFileFormat"/>; CSV has no signature to sniff, so it
    /// has no counterpart there and must always be requested explicitly.
    /// </summary>
    internal static class NativeFormat
    {
        internal const int Auto = 0;
        internal const int Xls = 1;
        internal const int Xlsx = 2;
        internal const int Xlsb = 3;
        internal const int Csv = 4;
    }
}
