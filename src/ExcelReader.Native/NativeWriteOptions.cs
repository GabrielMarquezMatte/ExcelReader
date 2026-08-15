using System.Buffers;
using System.Runtime.InteropServices;
using ExcelReader.Core.Writer;

namespace ExcelReader.Native
{
    /// <summary>
    /// Flat C ABI representation of <c>xl_write_options</c>. Numeric fields are 0 for "use the library
    /// default"; boolean-shaped fields use <see cref="NativeOptionState"/> for the same reason
    /// <see cref="NativeOpenOptionsRaw"/> does. See excelreader.h for the authoritative field list.
    /// </summary>
    [StructLayout(LayoutKind.Sequential)]
    internal unsafe struct NativeWriteOptionsRaw
    {
        public int StructSize;
        public int SheetNameLen;
        public byte* SheetName;

        public int CsvDelimiter;
        public int CsvQuote;
        public int Date1904;
        public int UseSharedStrings;
    }

    /// <summary>
    /// Decoded, pointer-free form of <see cref="NativeWriteOptionsRaw"/> — null means "use the library
    /// default", so only what the caller actually set is ever overridden.
    /// </summary>
    internal readonly struct NativeWriteOptions
    {
        /// <summary>Excel's own limit; a longer name is rejected here rather than by the writer, so the
        /// caller gets XL_INVALID_ARGUMENT before a file is created instead of XL_ERROR after.</summary>
        private const int MaxSheetNameLength = 31;

        // Excel's reserved sheet-name characters. Kept as a literal rather than reaching into Core:
        // IWorkbookWriter.AddSheet documents exactly this set, and the writer enforces it too.
        private const string ForbiddenSheetNameCharacters = @":\/?*[]";
        private static readonly SearchValues<char> ForbiddenSheetNameCharactersSearchValues = SearchValues.Create(ForbiddenSheetNameCharacters);
        internal string? SheetName { get; init; }
        internal byte? CsvDelimiter { get; init; }
        internal byte? CsvQuote { get; init; }
        internal bool? Date1904 { get; init; }
        internal bool? UseSharedStrings { get; init; }

        internal CsvWriterOptions ToCsvWriterOptions()
        {
            CsvWriterOptions options = CsvWriterOptions.Default;
            if (CsvDelimiter is byte delimiter)
            {
                options = options with { Delimiter = delimiter };
            }
            if (CsvQuote is byte quote)
            {
                options = options with { Quote = quote };
            }
            return options;
        }

        /// <summary>
        /// Validates and decodes a raw ABI struct. <paramref name="sheetName"/> arrives already
        /// UTF-8-decoded by <see cref="Exports"/>, since everything below that layer must stay
        /// pointer-free to remain testable.
        /// </summary>
        internal static bool TryDecode(NativeWriteOptionsRaw raw, string? sheetName, out NativeWriteOptions options, out string? error)
        {
            options = default;
            error = null;
            int expectedSize = Marshal.SizeOf<NativeWriteOptionsRaw>();
            if (raw.StructSize != expectedSize)
            {
                error = $"xl_write_options.struct_size is {raw.StructSize}, but this library expects {expectedSize}.";
                return false;
            }

            if (!TryValidateSheetName(sheetName, out error)
                || !TryDecodeByte(raw.CsvDelimiter, "csv_delimiter", out byte? delimiter, out error)
                || !TryDecodeByte(raw.CsvQuote, "csv_quote", out byte? quote, out error)
                || !TryDecodeState(raw.Date1904, "date1904", out bool? date1904, out error)
                || !TryDecodeState(raw.UseSharedStrings, "use_shared_strings", out bool? sharedStrings, out error))
            {
                return false;
            }

            options = new NativeWriteOptions
            {
                SheetName = sheetName,
                CsvDelimiter = delimiter,
                CsvQuote = quote,
                Date1904 = date1904,
                UseSharedStrings = sharedStrings,
            };
            return true;
        }

        private static bool TryValidateSheetName(string? sheetName, out string? error)
        {
            error = null;
            if (sheetName is null)
            {
                return true;
            }
            if (sheetName.Length is 0 or > MaxSheetNameLength)
            {
                error = $"xl_write_options.sheet_name must be 1-{MaxSheetNameLength} characters; got {sheetName.Length}.";
                return false;
            }
            if (sheetName.AsSpan().IndexOfAny(ForbiddenSheetNameCharactersSearchValues) >= 0)
            {
                error = $@"xl_write_options.sheet_name must not contain any of : \ / ? * [ ] ; got ""{sheetName}"".";
                return false;
            }
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
                error = $"xl_write_options.{fieldName} must be 0 (default) or a byte value 1-255; got {value}.";
                return false;
            }
            decoded = (byte)value;
            return true;
        }

        private static bool TryDecodeState(int value, string fieldName, out bool? decoded, out string? error)
        {
            decoded = null;
            error = null;
            if (value is not (NativeOptionState.Default or NativeOptionState.False or NativeOptionState.True))
            {
                error = $"xl_write_options.{fieldName} must be XL_OPT_DEFAULT/FALSE/TRUE (0/1/2); got {value}.";
                return false;
            }
            if (value != NativeOptionState.Default)
            {
                decoded = value == NativeOptionState.True;
            }
            return true;
        }
    }
}
