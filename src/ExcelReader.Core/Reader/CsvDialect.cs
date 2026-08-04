using System.Text;

namespace ExcelReader.Core.Reader
{
    /// <summary>The delimiter, quote, and encoding inferred by <see cref="CsvSniffer"/> from a sample of a delimited-text source.</summary>
    public readonly record struct CsvDialect
    {
        /// <summary>Gets the byte used to separate fields within a record.</summary>
        public byte Delimiter { get; init; }

        /// <summary>Gets the byte used to quote fields.</summary>
        public byte Quote { get; init; }

        /// <summary>Gets the encoding detected from a leading byte-order mark, or <see langword="null"/> when no byte-order mark was found (assume UTF-8).</summary>
        public Encoding? Encoding { get; init; }

        /// <summary>Gets a value indicating whether a byte-order mark was found at the start of the sample.</summary>
        public bool HasByteOrderMark { get; init; }

        /// <summary>Gets the dialect returned when a sample does not allow a delimiter to be determined: comma, double quote, UTF-8, no byte-order mark.</summary>
        public static CsvDialect Default { get; } = new()
        {
            Delimiter = (byte)',',
            Quote = (byte)'"',
            Encoding = null,
            HasByteOrderMark = false,
        };
    }
}
