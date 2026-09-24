using System.Buffers;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;
using ExcelReader.Core.Writer.Csv;

namespace ExcelReader.Native.Writer
{
    // NativeOpenOptionsRaw does. See excelreader.h for the authoritative field list.
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

    internal readonly struct NativeWriteOptions
    {
        private const string OptionsName = "xl_write_options";

        private const int MaxSheetNameLength = 31;

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

        internal static bool TryDecode(NativeWriteOptionsRaw raw, string? sheetName, out NativeWriteOptions options, out string? error)
        {
            options = default;
            if (!TryValidateStructSize(raw, out error))
            {
                return false;
            }

            if (!TryValidateSheetName(sheetName, out error)
                || !NativeOptionDecode.TryByte(raw.CsvDelimiter, OptionsName, "csv_delimiter", out byte? delimiter, out error)
                || !NativeOptionDecode.TryByte(raw.CsvQuote, OptionsName, "csv_quote", out byte? quote, out error)
                || !NativeOptionDecode.TryState(raw.Date1904, OptionsName, "date1904", out bool? date1904, out error)
                || !NativeOptionDecode.TryState(raw.UseSharedStrings, OptionsName, "use_shared_strings", out bool? sharedStrings, out error))
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

        internal static bool TryValidateStructSize(NativeWriteOptionsRaw raw, [NotNullWhen(false)] out string? error)
        {
            error = null;
            int expectedSize = Marshal.SizeOf<NativeWriteOptionsRaw>();
            if (raw.StructSize != expectedSize)
            {
                error = $"{OptionsName}.struct_size is {raw.StructSize}, but this library expects {expectedSize}.";
                return false;
            }
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
                error = $"{OptionsName}.sheet_name must be 1-{MaxSheetNameLength} characters; got {sheetName.Length}.";
                return false;
            }
            if (sheetName.AsSpan().IndexOfAny(ForbiddenSheetNameCharactersSearchValues) >= 0)
            {
                error = $@"{OptionsName}.sheet_name must not contain any of : \ / ? * [ ] ; got ""{sheetName}"".";
                return false;
            }
            return true;
        }
    }
}
