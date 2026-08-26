using ExcelReader.Core.Enums;

namespace ExcelReader.Core.Reader
{
    /// <summary>
    /// One column's guessed shape, as returned by
    /// <see cref="Excel.InferSchema(IExcelRowReader, int, int)"/>.
    /// </summary>
    /// <remarks>
    /// This is a guess over a bounded sample, not a guarantee about the whole sheet — verify it fits
    /// before trusting it. Feed it into <see cref="Parser.ExcelFluentParser{T}"/> to build a real map.
    /// </remarks>
    public readonly record struct ExcelColumnSchema
    {
        /// <summary>Gets the column's zero-based position in the sheet.</summary>
        public int Index { get; init; }

        /// <summary>
        /// Gets the column's trimmed header text, or <see langword="null"/> when the header row was
        /// skipped or that header cell was blank — in which case the column is addressable only by
        /// <see cref="Index"/>.
        /// </summary>
        public string? Name { get; init; }

        /// <summary>Gets the value type every sampled cell in this column agreed on.</summary>
        public ExcelColumnType Type { get; init; }

        /// <summary>Gets a value indicating whether any sampled row left this column empty.</summary>
        public bool IsNullable { get; init; }
    }
}
