using System.Globalization;
using Apache.Arrow;
using ExcelReader.Core.Reader;
using ExcelReader.Core.Reader.Columnar;
using ExcelReader.Core.Reader.Schema;

namespace ExcelReader.Arrow
{
    /// <summary>Converts an <see cref="IExcelSheet"/> into an Apache Arrow <see cref="RecordBatch"/>.</summary>
    public static class ArrowConversionExtensions
    {
        /// <summary>Converts one sheet into an Arrow <see cref="RecordBatch"/>.</summary>
        /// <param name="sheet">The sheet to convert.</param>
        /// <param name="options">The schema, header row and parallelism; <see cref="ArrowReadOptions.Default"/> when <see langword="null"/>.</param>
        /// <returns>A record batch with one column per schema entry and one row per data row.</returns>
        /// <exception cref="ArgumentNullException"><paramref name="sheet"/> is <see langword="null"/>.</exception>
        /// <exception cref="ArgumentOutOfRangeException"><see cref="ArrowReadOptions.DegreeOfParallelism"/> is negative.</exception>
        /// <exception cref="InvalidOperationException">
        /// A cell failed to convert to its column's declared type on a non-nullable column.
        /// </exception>
        /// <exception cref="ArgumentException">The sheet has fewer rows than <see cref="ArrowReadOptions.HeaderRow"/>.</exception>
        /// <exception cref="NotSupportedException"><see cref="ArrowReadOptions.Schema"/> names an unrecognized <see cref="ExcelColumnType"/>.</exception>
        public static RecordBatch ToArrowRecordBatch(this IExcelSheet sheet, ArrowReadOptions? options = null)
        {
            ArgumentNullException.ThrowIfNull(sheet);
            options ??= ArrowReadOptions.Default;
            ArgumentOutOfRangeException.ThrowIfNegative(options.DegreeOfParallelism, nameof(options));
            int headerRow = options.HeaderRow;
            IReadOnlyList<ExcelColumnSchema> schema = options.Schema ?? Excel.InferSchema(sheet, new SchemaInferenceOptions { HeaderRow = headerRow });
            TypedColumnSpec[] specs = ToSpecs(schema);
            using ColumnarTable table = TypedTableParser.Parse(sheet, specs, headerRow, options.DegreeOfParallelism, allowDictionary: true,
                chunkSizeOverride: 0, out _, out string? specError) ?? throw new ArgumentException(specError, nameof(options));

            Field[] fields = new Field[specs.Length];
            IArrowArray[] arrays = new IArrowArray[specs.Length];
            for (int i = 0; i < specs.Length; i++)
            {
                fields[i] = new Field(specs[i].Label, ArrowColumns.TypeOf(specs[i].Type), specs[i].Nullable);
                arrays[i] = ArrowColumns.ToArray(in table.Columns[i]);
            }
            return new RecordBatch(new Schema(fields, metadata: null), arrays, checked((int)table.RowCount));
        }

        private static TypedColumnSpec[] ToSpecs(IReadOnlyList<ExcelColumnSchema> schema)
        {
            TypedColumnSpec[] specs = new TypedColumnSpec[schema.Count];
            for (int i = 0; i < schema.Count; i++)
            {
                ExcelColumnSchema column = schema[i];
                if (!Enum.IsDefined(column.Type))
                {
                    throw new NotSupportedException($"column type {column.Type} is not supported yet.");
                }
                specs[i] = new TypedColumnSpec
                {
                    Index = column.Index,
                    Type = column.Type,
                    Nullable = column.IsNullable,
                    Label = column.Name ?? column.Index.ToString(CultureInfo.InvariantCulture),
                };
            }
            return specs;
        }
    }
}
