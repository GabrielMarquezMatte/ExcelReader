using ExcelReader.Core.Reader.Schema;

namespace ExcelReader.Arrow
{
    /// <summary>Options for converting a sheet with <see cref="ArrowConversionExtensions.ToArrowRecordBatch"/>.</summary>
    public sealed class ArrowReadOptions
    {
        /// <summary>Gets the default options: header on row 1, schema inferred from the sheet.</summary>
        public static ArrowReadOptions Default { get; } = new();

        /// <summary>
        /// Gets the column schema to use. <see langword="null"/>, the default, infers it from the sheet with
        /// <see cref="Core.Reader.Excel.InferSchema"/> using <see cref="HeaderRow"/>.
        /// </summary>
        public IReadOnlyList<ExcelColumnSchema>? Schema { get; init; }

        /// <summary>
        /// Gets the 1-based header row, used both for schema inference (when <see cref="Schema"/> is
        /// <see langword="null"/>) and to skip the header line during the data pass. <c>0</c> means no header
        /// row. Defaults to <c>1</c>.
        /// </summary>
        public int HeaderRow { get; init; } = 1;
    }
}
