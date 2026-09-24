using System.Text;

namespace ExcelReader.Core.Writer.Csv
{
    /// <summary>Configures the dialect, encoding, and record terminator used when writing a CSV file.</summary>
    /// <remarks>Mirrors <see cref="Reader.Csv.CsvReaderOptions"/> property for property, so a file written with
    /// one set of settings reads back with the matching set.</remarks>
    public sealed record CsvWriterOptions
    {
        /// <summary>The byte written between fields. Defaults to <c>,</c>.</summary>
        public byte Delimiter { get; init; } = (byte)',';

        /// <summary>The byte used to quote a field that contains the delimiter, itself, or a line break. Defaults to <c>"</c>.</summary>
        public byte Quote { get; init; } = (byte)'"';

        /// <summary>The text encoding to write. When not <see langword="null"/>, output is transcoded from UTF-8
        /// as it is written. Defaults to <see langword="null"/>, meaning UTF-8 is written directly.</summary>
        /// <remarks>Counterpart to <see cref="Reader.Csv.CsvReaderOptions.Encoding"/>: fields are always formatted
        /// as UTF-8 internally, so a non-UTF-8 target is transcoded on the way out.</remarks>
        public Encoding? Encoding { get; init; }

        /// <summary>Whether the encoding's byte-order mark is written before the first record. Defaults to
        /// <see langword="false"/>.</summary>
        /// <remarks>Counterpart to <see cref="Reader.Csv.CsvReaderOptions.DetectEncodingFromByteOrderMark"/>. Writing
        /// one is what lets Excel open a UTF-8 file as UTF-8 rather than as the system code page.</remarks>
        public bool WriteByteOrderMark { get; init; }

        /// <summary>The byte sequence written after each record. Defaults to <see cref="CsvNewLine.CarriageReturnLineFeed"/>,
        /// which is what RFC 4180 specifies.</summary>
        /// <remarks>Fields containing <c>\r</c> or <c>\n</c> are quoted either way, so the choice does not change
        /// how the file reads back.</remarks>
        public CsvNewLine NewLine { get; init; } = CsvNewLine.CarriageReturnLineFeed;

        /// <summary>The default options: comma-delimited, double-quote-quoted, UTF-8 with no byte-order mark, CRLF-terminated.</summary>
        public static CsvWriterOptions Default { get; } = new();
    }
}
