using System.Text;

namespace ExcelReader.Core.Reader
{
    /// <summary>Options controlling how <see cref="CsvReader"/> parses a CSV source.</summary>
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

        /// <summary>Gets the maximum byte length allowed for a single buffered record/field. Defaults to 32 MiB.</summary>
        /// <remarks>Mirrors <see cref="ExcelReaderOptions.MaxCellBytes"/> for the Excel readers.</remarks>
        public int MaxCellBytes { get; init; } = 32 * 1024 * 1024;

        /// <summary>Gets the default options instance, used whenever a <see cref="CsvReader"/> is opened without explicit options.</summary>
        public static CsvReaderOptions Default { get; } = new();
    }
}
