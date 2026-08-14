using System.Runtime.InteropServices;
using ExcelReader.Core.Reader;

namespace ExcelReader.Native
{
    /// <summary>Boolean-shaped option states. "0" is never ambiguous between "off" and "use the library
    /// default" — several <see cref="NativeOpenOptionsRaw"/> fields default to true. Mirrors
    /// XL_OPT_* in include/excelreader.h.</summary>
    internal static class NativeOptionState
    {
        internal const int Default = 0;
        internal const int False = 1;
        internal const int True = 2;
    }

    /// <summary>
    /// Flat C ABI representation of <c>xl_open_options</c>. Every numeric field is 0 for "use the
    /// library default"; every boolean-shaped field uses <see cref="NativeOptionState"/> instead of a
    /// plain 0/1, since several of them default to true. See excelreader.h for the authoritative field
    /// list and comments.
    /// </summary>
    [StructLayout(LayoutKind.Sequential)]
    internal struct NativeOpenOptionsRaw
    {
        public int StructSize;

        public int CsvSniffDialect;
        public int CsvDelimiter;
        public int CsvQuote;
        public int CsvDetectBom;
        public int CsvMaxCellBytes;
        public int CsvInternStrings;

        public long MaxTotalDecompressedBytes;
        public int MaxCellBytes;
        public long MaxSharedStringBytes;
        public int MaxZipEntries;
        public int PrefetchDecompression;
        public int InternStrings;
    }

    /// <summary>
    /// Decoded, validated form of <see cref="NativeOpenOptionsRaw"/> — every field is nullable, with
    /// null meaning "use the library default", so <see cref="ToCsvReaderOptions"/>/
    /// <see cref="ToExcelReaderOptions"/> only ever override what the caller actually set.
    /// </summary>
    internal readonly struct NativeOpenOptions
    {
        internal bool CsvSniffDialect { get; init; }
        internal byte? CsvDelimiter { get; init; }
        internal byte? CsvQuote { get; init; }
        internal bool? CsvDetectByteOrderMark { get; init; }
        internal int? CsvMaxCellBytes { get; init; }
        internal bool? CsvInternStrings { get; init; }

        internal long? MaxTotalDecompressedBytes { get; init; }
        internal int? MaxCellBytes { get; init; }
        internal long? MaxSharedStringBytes { get; init; }
        internal int? MaxZipEntries { get; init; }
        internal bool? PrefetchDecompression { get; init; }
        internal bool? InternStrings { get; init; }

        internal CsvReaderOptions ToCsvReaderOptions()
        {
            CsvReaderOptions options = CsvReaderOptions.Default;
            if (CsvDelimiter is byte delimiter)
            {
                options = options with { Delimiter = delimiter };
            }
            if (CsvQuote is byte quote)
            {
                options = options with { Quote = quote };
            }
            if (CsvDetectByteOrderMark is bool detectBom)
            {
                options = options with { DetectEncodingFromByteOrderMark = detectBom };
            }
            if (CsvMaxCellBytes is int maxCellBytes)
            {
                options = options with { MaxCellBytes = maxCellBytes };
            }
            if (CsvInternStrings is bool internStrings)
            {
                options = options with { InternStrings = internStrings };
            }
            return options;
        }

        internal ExcelReaderOptions ToExcelReaderOptions()
        {
            ExcelReaderOptions options = ExcelReaderOptions.Default;
            if (MaxTotalDecompressedBytes is long maxTotal)
            {
                options = options with { MaxTotalDecompressedBytes = maxTotal };
            }
            if (MaxCellBytes is int maxCellBytes)
            {
                options = options with { MaxCellBytes = maxCellBytes };
            }
            if (MaxSharedStringBytes is long maxSharedStrings)
            {
                options = options with { MaxSharedStringBytes = maxSharedStrings };
            }
            if (MaxZipEntries is int maxZipEntries)
            {
                options = options with { MaxZipEntries = maxZipEntries };
            }
            if (PrefetchDecompression is bool prefetch)
            {
                options = options with { PrefetchDecompression = prefetch };
            }
            if (InternStrings is bool internStrings)
            {
                options = options with { InternStrings = internStrings };
            }
            return options;
        }

        /// <summary>
        /// Validates and decodes a raw ABI struct. Returns <see langword="false"/> (with no exception —
        /// callers are on the hot "was this argument well-formed" path, not an error-recovery one) for an
        /// unrecognized <see cref="NativeOpenOptionsRaw.StructSize"/> or an out-of-range field.
        /// </summary>
        internal static bool TryDecode(NativeOpenOptionsRaw raw, out NativeOpenOptions options, out string? error)
        {
            options = default;
            error = null;
            int expectedSize = Marshal.SizeOf<NativeOpenOptionsRaw>();
            if (raw.StructSize != expectedSize)
            {
                error = $"xl_open_options.struct_size is {raw.StructSize}, but this library expects {expectedSize}.";
                return false;
            }

            if (!TryDecodeByte(raw.CsvDelimiter, "csv_delimiter", out byte? delimiter, out error)
                || !TryDecodeByte(raw.CsvQuote, "csv_quote", out byte? quote, out error)
                || !TryDecodeNonNegative(raw.CsvMaxCellBytes, "csv_max_cell_bytes", out int? csvMaxCellBytes, out error)
                || !TryDecodeNonNegative(raw.MaxCellBytes, "max_cell_bytes", out int? maxCellBytes, out error)
                || !TryDecodeNonNegative(raw.MaxZipEntries, "max_zip_entries", out int? maxZipEntries, out error)
                || !TryDecodeNonNegativeLong(raw.MaxTotalDecompressedBytes, "max_total_decompressed_bytes", out long? maxTotal, out error)
                || !TryDecodeNonNegativeLong(raw.MaxSharedStringBytes, "max_shared_string_bytes", out long? maxSharedStrings, out error))
            {
                return false;
            }

            if (!TryDecodeState(raw.CsvSniffDialect, "csv_sniff_dialect", out bool? sniffDialect, out error)
                || !TryDecodeState(raw.CsvDetectBom, "csv_detect_bom", out bool? detectBom, out error)
                || !TryDecodeState(raw.CsvInternStrings, "csv_intern_strings", out bool? csvInternStrings, out error)
                || !TryDecodeState(raw.PrefetchDecompression, "prefetch_decompression", out bool? prefetch, out error)
                || !TryDecodeState(raw.InternStrings, "intern_strings", out bool? internStrings, out error))
            {
                return false;
            }

            options = new NativeOpenOptions
            {
                CsvSniffDialect = sniffDialect ?? false,
                CsvDelimiter = delimiter,
                CsvQuote = quote,
                CsvDetectByteOrderMark = detectBom,
                CsvMaxCellBytes = csvMaxCellBytes,
                CsvInternStrings = csvInternStrings,
                MaxTotalDecompressedBytes = maxTotal,
                MaxCellBytes = maxCellBytes,
                MaxSharedStringBytes = maxSharedStrings,
                MaxZipEntries = maxZipEntries,
                PrefetchDecompression = prefetch,
                InternStrings = internStrings,
            };
            return true;
        }

        private static bool TryDecodeByte(int value, string fieldName, out byte? decoded, out string? error)
        {
            decoded = null;
            error = null;
            if (value == 0)
            {
                return true;
            }
            if (value is < 1 or > 255)
            {
                error = $"xl_open_options.{fieldName} must be 0 (default) or a byte value 1-255; got {value}.";
                return false;
            }
            decoded = (byte)value;
            return true;
        }

        private static bool TryDecodeNonNegative(int value, string fieldName, out int? decoded, out string? error)
        {
            decoded = null;
            error = null;
            if (value == 0)
            {
                return true;
            }
            if (value < 0)
            {
                error = $"xl_open_options.{fieldName} must be 0 (default) or a positive value; got {value}.";
                return false;
            }
            decoded = value;
            return true;
        }

        private static bool TryDecodeNonNegativeLong(long value, string fieldName, out long? decoded, out string? error)
        {
            decoded = null;
            error = null;
            if (value == 0)
            {
                return true;
            }
            if (value < 0)
            {
                error = $"xl_open_options.{fieldName} must be 0 (default) or a positive value; got {value}.";
                return false;
            }
            decoded = value;
            return true;
        }

        private static bool TryDecodeState(int value, string fieldName, out bool? decoded, out string? error)
        {
            decoded = value switch
            {
                NativeOptionState.Default => null,
                NativeOptionState.False => false,
                NativeOptionState.True => true,
                _ => null,
            };
            error = null;
            if (value is not (NativeOptionState.Default or NativeOptionState.False or NativeOptionState.True))
            {
                error = $"xl_open_options.{fieldName} must be XL_OPT_DEFAULT/FALSE/TRUE (0/1/2); got {value}.";
                return false;
            }
            return true;
        }
    }
}
