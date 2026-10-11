using System.Text;

namespace ExcelReader.Core.Reader.Csv
{
    /// <summary>The delimited-text settings for a CSV source, carried by <see cref="ExcelReaderOptions.Csv"/>.</summary>
    /// <remarks>Size limits and string interning come from <see cref="ExcelReaderOptions"/>, as for every other format.</remarks>
    public sealed record CsvReaderOptions
    {
        /// <summary>Gets the byte used to separate fields within a record. Defaults to <c>,</c>.</summary>
        public byte Delimiter { get; init; } = (byte)',';

        /// <summary>Gets the byte used to quote fields. Defaults to <c>"</c>.</summary>
        public byte Quote { get; init; } = (byte)'"';

        /// <summary>Gets the source's text encoding. When not <see langword="null"/>, the source is transcoded to UTF-8 as it is read. Defaults to <see langword="null"/>, meaning the source is already UTF-8.</summary>
        /// <remarks>
        /// Non-UTF-8 sources (e.g. Windows-1252 exports) are transcoded at open time, so the scanner
        /// itself always works in UTF-8.
        /// </remarks>
        public Encoding? Encoding { get; init; }

        /// <summary>Gets a value indicating whether a leading UTF-8 byte-order mark should be detected and skipped. Defaults to <see langword="true"/>.</summary>
        public bool DetectEncodingFromByteOrderMark { get; init; } = true;

        internal int MaxCellBytes { get; init; } = 32 * 1024 * 1024;

        internal bool InternStrings { get; init; }

        // ponytail: the native ABI keeps separate csv_max_cell_bytes/csv_intern_strings fields; these let it win over ExcelReaderOptions.
        internal int? MaxCellBytesOverride { get; init; }

        internal bool? InternStringsOverride { get; init; }

        /// <summary>
        /// Gets a value indicating whether the source's dialect is inferred from a sample of its leading
        /// bytes at open time, replacing <see cref="Delimiter"/> and <see cref="Quote"/> with what
        /// <see cref="CsvSniffer"/> reports. Defaults to <see langword="false"/>.
        /// </summary>
        /// <remarks>
        /// Applied by the entry points that own the source — <see cref="Excel.FromCsvFile"/> and the other
        /// <c>Excel.FromCsv*</c> factories, <see cref="Excel.Open(string,ExcelFileFormat,ExcelReaderOptions?)"/>
        /// with <see cref="ExcelFileFormat.Csv"/>, and the parallel CSV entry points — since the sample
        /// has to be read before the reader starts. An explicitly set <see cref="Encoding"/> is kept unless the
        /// source carries a byte-order mark naming a different one. Sniffing a stream reads a sample and
        /// restores the position, so it requires a seekable stream.
        /// </remarks>
        public bool SniffDialect { get; init; }

        /// <summary>Gets the default options instance, used whenever a <see cref="CsvReader"/> is opened without explicit options.</summary>
        public static CsvReaderOptions Default { get; } = new();

        /// <summary>Returns a copy of these options with the delimiter, quote, and encoding of the given dialect.</summary>
        /// <param name="dialect">The dialect to apply, typically produced by <see cref="CsvSniffer.Detect(ReadOnlySpan{byte}, CsvSnifferOptions?)"/>.</param>
        public CsvReaderOptions WithDialect(CsvDialect dialect)
        {
            return this with
            {
                Delimiter = dialect.Delimiter,
                Quote = dialect.Quote,
                Encoding = dialect.Encoding,
            };
        }
    }
}
