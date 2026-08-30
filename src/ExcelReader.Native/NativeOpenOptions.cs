using System.Numerics;
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
    /// Field-level decoding shared by <see cref="NativeOpenOptions"/> and
    /// <see cref="NativeWriteOptions"/>. Both option structs carry the same two shapes of field — a
    /// byte-valued one and a <see cref="NativeOptionState"/> tri-state — with the same rules and the
    /// same messages; only the struct's own name differs, so it arrives as an argument.
    /// </summary>
    internal static class NativeOptionDecode
    {
        /// <summary>A byte-valued field: 0 means "use the library default", 1-255 is a real byte.</summary>
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

        /// <summary>A boolean-shaped field, encoded as <see cref="NativeOptionState"/> rather than a plain
        /// 0/1 because several of these default to true.</summary>
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

        // Appended at the end, never inserted in the middle — see StructSize's exact-equality check
        // in TryDecode, which is what turns a stale caller's layout mismatch into a loud
        // XL_INVALID_ARGUMENT instead of silently misreading these two fields (or everything after them).
        public IntPtr Password;
        public int PasswordLen;
    }

    /// <summary>
    /// Decoded, validated form of <see cref="NativeOpenOptionsRaw"/> — every field is nullable, with
    /// null meaning "use the library default", so <see cref="ToCsvReaderOptions"/>/
    /// <see cref="ToExcelReaderOptions"/> only ever override what the caller actually set.
    /// </summary>
    internal readonly struct NativeOpenOptions
    {
        /// <summary>The C struct's name, as it appears in every message this type produces.</summary>
        private const string OptionsName = "xl_open_options";

        // An unbounded password length is the same class of hole as an unbounded count arriving as an
        // argument (see NativeLimits' remarks) — it just arrives here instead.
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

        /// <summary>Password for an encrypted OOXML workbook, decoded from the raw pointer+length pair.
        /// Null means "no password supplied" — either the workbook isn't encrypted, or the caller wants
        /// the library-default behavior of failing with <see cref="ExcelEncryptionReason.PasswordRequired"/>.</summary>
        internal string? Password { get; init; }

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
            if (Password is not null)
            {
                options = options with { Password = Password };
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

            // UTF-8 with an explicit length rather than NUL-terminated: a password may contain anything.
            // The pointer need only be valid for this call; we copy immediately.
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

        // Serves both the int and long fields: the rule ("0 means default, negative is a caller error")
        // and its message are identical, and only the width differed.
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
