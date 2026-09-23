using System.Numerics;
using System.Runtime.InteropServices;
using ExcelReader.Core.Reader;

namespace ExcelReader.Native
{
    internal static class NativeOptionState
    {
        internal const int Default = 0;
        internal const int False = 1;
        internal const int True = 2;
    }

    internal static class NativeOptionDecode
    {
        internal static bool TryByte(int value, string structName, string fieldName, out byte? decoded, out string? error)
        {
            decoded = null;
            error = null;
            if (value == 0)
            {
                return true;
            }
            if (value is < 1 or > 255)
            {
                error = $"{structName}.{fieldName} must be 0 (default) or a byte value 1-255; got {value}.";
                return false;
            }
            decoded = (byte)value;
            return true;
        }

        internal static bool TryState(int value, string structName, string fieldName, out bool? decoded, out string? error)
        {
            decoded = null;
            error = null;
            if (value is not (NativeOptionState.Default or NativeOptionState.False or NativeOptionState.True))
            {
                error = $"{structName}.{fieldName} must be XL_OPT_DEFAULT/FALSE/TRUE (0/1/2); got {value}.";
                return false;
            }
            if (value != NativeOptionState.Default)
            {
                decoded = value == NativeOptionState.True;
            }
            return true;
        }
    }

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

        public IntPtr Password;
        public int PasswordLen;
    }

    internal readonly struct NativeOpenOptions
    {
        private const string OptionsName = "xl_open_options";

        private const int MaxPasswordBytes = 4096;

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

        internal string? Password { get; init; }

        internal CsvReaderOptions ToCsvReaderOptions()
        {
            CsvReaderOptions options = CsvSniffDialect ? CsvReaderOptions.Default with { SniffDialect = true } : CsvReaderOptions.Default;
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
            if (Password is not null)
            {
                options = options with { Password = Password };
            }
            return options with { Csv = ToCsvReaderOptions() };
        }

        internal static bool TryDecode(NativeOpenOptionsRaw raw, out NativeOpenOptions options, out string? error)
        {
            options = default;
            error = null;
            int expectedSize = Marshal.SizeOf<NativeOpenOptionsRaw>();
            if (raw.StructSize != expectedSize)
            {
                error = $"{OptionsName}.struct_size is {raw.StructSize}, but this library expects {expectedSize}.";
                return false;
            }

            if (!NativeOptionDecode.TryByte(raw.CsvDelimiter, OptionsName, "csv_delimiter", out byte? delimiter, out error)
                || !NativeOptionDecode.TryByte(raw.CsvQuote, OptionsName, "csv_quote", out byte? quote, out error)
                || !TryDecodeNonNegative(raw.CsvMaxCellBytes, "csv_max_cell_bytes", out int? csvMaxCellBytes, out error)
                || !TryDecodeNonNegative(raw.MaxCellBytes, "max_cell_bytes", out int? maxCellBytes, out error)
                || !TryDecodeNonNegative(raw.MaxZipEntries, "max_zip_entries", out int? maxZipEntries, out error)
                || !TryDecodeNonNegative(raw.MaxTotalDecompressedBytes, "max_total_decompressed_bytes", out long? maxTotal, out error)
                || !TryDecodeNonNegative(raw.MaxSharedStringBytes, "max_shared_string_bytes", out long? maxSharedStrings, out error))
            {
                return false;
            }

            if (!NativeOptionDecode.TryState(raw.CsvSniffDialect, OptionsName, "csv_sniff_dialect", out bool? sniffDialect, out error)
                || !NativeOptionDecode.TryState(raw.CsvDetectBom, OptionsName, "csv_detect_bom", out bool? detectBom, out error)
                || !NativeOptionDecode.TryState(raw.CsvInternStrings, OptionsName, "csv_intern_strings", out bool? csvInternStrings, out error)
                || !NativeOptionDecode.TryState(raw.PrefetchDecompression, OptionsName, "prefetch_decompression", out bool? prefetch, out error)
                || !NativeOptionDecode.TryState(raw.InternStrings, OptionsName, "intern_strings", out bool? internStrings, out error))
            {
                return false;
            }

            string? password = null;
            if (raw.Password != IntPtr.Zero)
            {
                if (raw.PasswordLen is < 0 or > MaxPasswordBytes)
                {
                    error = $"{OptionsName}.password_len must be between 0 and {MaxPasswordBytes}; got {raw.PasswordLen}.";
                    return false;
                }
                unsafe
                {
                    password = System.Text.Encoding.UTF8.GetString((byte*)raw.Password, raw.PasswordLen);
                }
            }
            else if (raw.PasswordLen != 0)
            {
                error = $"{OptionsName}.password_len is {raw.PasswordLen} but password is null.";
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
                Password = password,
            };
            return true;
        }

        private static bool TryDecodeNonNegative<T>(T value, string fieldName, out T? decoded, out string? error)
            where T : struct, INumberBase<T>
        {
            decoded = null;
            error = null;
            if (T.IsZero(value))
            {
                return true;
            }
            if (T.IsNegative(value))
            {
                error = $"xl_open_options.{fieldName} must be 0 (default) or a positive value; got {value}.";
                return false;
            }
            decoded = value;
            return true;
        }
    }
}
