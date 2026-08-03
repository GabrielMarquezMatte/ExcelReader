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

        /// <summary>
        /// Gets a value indicating whether repeated field text is deduplicated through a small
        /// content-keyed cache when materialized via <see cref="ValueObjects.Cell.GetString"/>.
        /// Defaults to <see langword="false"/>. Has no effect on the zero-copy
        /// <see cref="ValueObjects.Cell.Value"/> span path — only <c>GetString()</c> consults the
        /// cache. Worth enabling for genuinely categorical columns (low-cardinality repeated text);
        /// for high-cardinality data (unique IDs, or numeric/date columns read back as text, which CSV
        /// has no other representation for) every lookup is a cache miss, so the per-call hashing cost
        /// is pure overhead with no dedup benefit — measured roughly doubling wall-clock time on a
        /// real-world, mixed-cardinality CSV corpus. Off by default so this never regresses a caller
        /// who hasn't measured their own data.
        /// </summary>
        public bool InternStrings { get; init; }

        /// <summary>Gets the default options instance, used whenever a <see cref="CsvReader"/> is opened without explicit options.</summary>
        public static CsvReaderOptions Default { get; } = new();
    }
}
