using System.Buffers;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;
using ExcelReader.Core.Writer;
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

        public int DegreeOfParallelism;
        public int BatchRows;

        internal const int V7Size = 32;

        internal static unsafe bool TryRead(NativeWriteOptionsRaw* options, out NativeWriteOptionsRaw raw, [NotNullWhen(false)] out string? error)
        {
            raw = default;
            error = null;
            if (options is null)
            {
                raw = new NativeWriteOptionsRaw { StructSize = sizeof(NativeWriteOptionsRaw), DegreeOfParallelism = 1 };
                return true;
            }
            int size = options->StructSize;
            if (size != sizeof(NativeWriteOptionsRaw) && size != V7Size)
            {
                error = $"xl_write_options.struct_size is {size}, but this library expects {sizeof(NativeWriteOptionsRaw)} (or {V7Size} for the 7.0 layout).";
                return false;
            }
            NativeWriteOptionsRaw copy = default;
            Buffer.MemoryCopy(options, &copy, sizeof(NativeWriteOptionsRaw), size);
            copy.StructSize = sizeof(NativeWriteOptionsRaw);
            if (size == V7Size)
            {
                copy.DegreeOfParallelism = 1;
            }
            raw = copy;
            return true;
        }
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
        internal int DegreeOfParallelism { get; init; }
        internal int BatchRows { get; init; }

        internal ParallelWriteOptions? ToParallelWriteOptions(int format)
        {
            if (format is not (NativeFormat.Xlsx or NativeFormat.Xlsb) || DegreeOfParallelism == 1)
            {
                return null;
            }
            return new ParallelWriteOptions
            {
                DegreeOfParallelism = DegreeOfParallelism,
                BatchRows = BatchRows == 0 ? ParallelWriteOptions.Default.BatchRows : BatchRows,
            };
        }

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
            if (raw.DegreeOfParallelism < 0)
            {
                error = $"{OptionsName}.degree_of_parallelism must be 0 (processor count) or positive; got {raw.DegreeOfParallelism}.";
                return false;
            }
            if (raw.BatchRows < 0)
            {
                error = $"{OptionsName}.batch_rows must be 0 (library default) or positive; got {raw.BatchRows}.";
                return false;
            }

            options = new NativeWriteOptions
            {
                SheetName = sheetName,
                CsvDelimiter = delimiter,
                CsvQuote = quote,
                Date1904 = date1904,
                UseSharedStrings = sharedStrings,
                DegreeOfParallelism = raw.DegreeOfParallelism,
                BatchRows = raw.BatchRows,
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
