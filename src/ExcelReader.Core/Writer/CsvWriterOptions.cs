namespace ExcelReader.Core.Writer
{
    /// <summary>Configures the field delimiter and quote character used when writing a CSV file.</summary>
    public sealed record CsvWriterOptions
    {
        /// <summary>The byte written between fields. Defaults to <c>,</c>.</summary>
        public byte Delimiter { get; init; } = (byte)',';

        /// <summary>The byte used to quote a field that contains the delimiter, itself, or a line break. Defaults to <c>"</c>.</summary>
        public byte Quote { get; init; } = (byte)'"';

        /// <summary>The default options: comma-delimited, double-quote-quoted.</summary>
        public static CsvWriterOptions Default { get; } = new();
    }
}
