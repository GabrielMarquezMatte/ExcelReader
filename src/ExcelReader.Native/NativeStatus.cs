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
    /// Upper bounds on the counts a caller hands across the ABI. Mirrors XL_MAX_* in
    /// include/excelreader.h.
    /// </summary>
    /// <remarks>
    /// Every one of these drives an allocation or a pointer walk over caller memory, so an
    /// unvalidated value is the same class of hole the readers guard against for file bytes (see
    /// STYLEGUIDE.md, "Untrusted Input") — the difference is only that the hostile number arrives as
    /// an argument instead. The values come from Excel's own ceilings, which
    /// <c>ExcelReader.Core.Reader.ExcelLimits</c> holds for the managed side: a request naming more
    /// columns than a sheet can hold, or a header wider than a cell can hold, cannot describe a real
    /// workbook and is rejected rather than clamped.
    /// </remarks>
    internal static class NativeLimits
    {
        /// <summary>A..XFD — no sheet has more columns, so no spec list needs more entries.</summary>
        internal const int MaxColumnSpecs = 16_384;

        /// <summary>Excel's 32,767-character cell limit at UTF-8's 4-byte worst case.</summary>
        internal const int MaxColumnNameBytes = 131_068;
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
