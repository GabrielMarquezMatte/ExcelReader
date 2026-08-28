using Apache.Arrow;
using ExcelReader.Core.Reader;
using ExcelReader.Core.ValueObjects;

namespace ExcelReader.Arrow
{
    /// <summary>Converts an <see cref="IExcelRowReader"/>'s current sheet into an Apache Arrow <see cref="RecordBatch"/>.</summary>
    public static class ArrowConversionExtensions
    {
        /// <summary>
        /// Reads <paramref name="reader"/>'s current sheet into one <see cref="RecordBatch"/>, entirely
        /// in memory.
        /// </summary>
        /// <param name="reader">The reader whose current sheet is converted.</param>
        /// <param name="schema">
        /// The column shape to convert to. When <see langword="null"/>, resolved via
        /// <see cref="Excel.InferSchema(IExcelRowReader, int, int)"/> using <paramref name="headerRow"/>.
        /// </param>
        /// <param name="headerRow">
        /// 1-based header row, reused both for schema inference (when <paramref name="schema"/> is
        /// <see langword="null"/>) and to skip the header line during the data pass. 0 means no header row.
        /// </param>
        /// <exception cref="ArgumentNullException"><paramref name="reader"/> is <see langword="null"/>.</exception>
        /// <exception cref="InvalidOperationException">
        /// A cell failed to convert to its column's declared type on a non-nullable column.
        /// </exception>
        public static RecordBatch ToArrowRecordBatch(this IExcelRowReader reader, ExcelColumnSchema[]? schema = null, int headerRow = 1)
        {
            ArgumentNullException.ThrowIfNull(reader);
            ExcelColumnSchema[] resolvedSchema = schema ?? Excel.InferSchema(reader, headerRow);

            ColumnAppender[] appenders = new ColumnAppender[resolvedSchema.Length];
            for (int i = 0; i < resolvedSchema.Length; i++)
            {
                appenders[i] = ColumnAppender.Create(resolvedSchema[i]);
            }

            using IExcelRowEnumerator rows = reader.GetEnumerator();
            SkipHeaderRow(rows, headerRow);

            int rowCount = 0;
            while (rows.MoveNext())
            {
                Row row = rows.Current;
                for (int i = 0; i < appenders.Length; i++)
                {
                    appenders[i].Append(row[resolvedSchema[i].Index], reader.IsDate1904);
                }
                rowCount++;
            }

            Field[] fields = new Field[appenders.Length];
            IArrowArray[] arrays = new IArrowArray[appenders.Length];
            for (int i = 0; i < appenders.Length; i++)
            {
                fields[i] = appenders[i].Field;
                arrays[i] = appenders[i].Build();
            }

            Schema arrowSchema = new(fields, metadata: null);
            return new RecordBatch(arrowSchema, arrays, rowCount);
        }

        private static void SkipHeaderRow(IExcelRowEnumerator rows, int headerRow)
        {
            for (int rowNumber = 1; rowNumber <= headerRow; rowNumber++)
            {
                if (!rows.MoveNext())
                {
                    throw new ArgumentException($"sheet has fewer than {headerRow} row(s); cannot resolve header_row.", nameof(headerRow));
                }
            }
        }
    }
}
