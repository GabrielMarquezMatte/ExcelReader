namespace ExcelReader.Core.Reader
{
    /// <summary>
    /// Resource limits and behavior toggles applied when opening and reading an Excel workbook (XLSX/XLSB/XLS).
    /// Exceeding a byte or entry-count limit throws <see cref="ExcelLimitExceededException"/>.
    /// </summary>
    public sealed record ExcelReaderOptions
    {
        /// <summary>Gets the maximum total decompressed bytes allowed across the whole workbook. Defaults to 512 MiB.</summary>
        /// <remarks>Applies to ZIP-backed formats (XLSX/XLSB) as their decompressed byte budget, and to
        /// the legacy CFB container (.xls) as the cap on its declared Workbook stream size — the CFB path
        /// has nothing to decompress, but this is still the caller's budget for what that phase may
        /// materialize.</remarks>
        public long MaxTotalDecompressedBytes { get; init; } = 512L * 1024 * 1024;

        /// <summary>Gets the maximum byte length allowed for a single cell's value. Defaults to 32 MiB.</summary>
        public int MaxCellBytes { get; init; } = 32 * 1024 * 1024;

        /// <summary>Gets the maximum total byte size allowed for the shared string table. Defaults to 128 MiB.</summary>
        public long MaxSharedStringBytes { get; init; } = 128L * 1024 * 1024;

        /// <summary>Gets the maximum number of ZIP entries allowed in an XLSX/XLSB archive. Defaults to 65,536.</summary>
        public int MaxZipEntries { get; init; } = 65_536;

        /// <summary>Gets a value indicating whether ZIP-backed sheet data is decompressed on a
        /// background thread ahead of parsing. Defaults to <see langword="false"/>.</summary>
        /// <remarks>Applies only to ZIP-backed formats (XLSX/XLSB); XLS and CSV have nothing to
        /// decompress, so the option is silently ignored for them. Intended for single-file batch
        /// processing, where overlapping inflate with parsing shortens one read's wall-clock time —
        /// not for concurrent server workloads, where the extra background thread per read competes
        /// with work already saturating the CPU.</remarks>
        public bool PrefetchDecompression { get; init; }

        /// <summary>Gets the default options instance, used whenever a reader is opened without explicit options.</summary>
        public static ExcelReaderOptions Default { get; } = new();
    }
}
