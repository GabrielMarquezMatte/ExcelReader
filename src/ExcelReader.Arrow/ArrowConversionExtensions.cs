using Apache.Arrow;
using ExcelReader.Core.Reader;
using ExcelReader.Core.Reader.Schema;

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
        /// <exception cref="ArgumentException">The sheet has fewer rows than <paramref name="headerRow"/>.</exception>
        /// <exception cref="NotSupportedException"><paramref name="schema"/> names an unrecognized <see cref="ExcelColumnType"/>.</exception>
        public static RecordBatch ToArrowRecordBatch(this IExcelRowReader reader, ExcelColumnSchema[]? schema = null, int headerRow = 1)
        {
            ArgumentNullException.ThrowIfNull(reader);
            ExcelColumnSchema[] resolvedSchema = schema ?? Excel.InferSchema(reader, headerRow);
            using IExcelRowEnumerator rows = reader.GetEnumerator();
            return Build(rows, reader.IsDate1904, resolvedSchema, headerRow);
        }

        /// <summary>Converts one sheet into an Arrow <see cref="RecordBatch"/>.</summary>
        /// <param name="sheet">The sheet to convert.</param>
        /// <param name="schema">The column schema to use; inferred from the sheet with <see cref="Excel.InferSchema(IExcelSheet, int, int, bool)"/> when <see langword="null"/>.</param>
        /// <param name="headerRow">1-based row number holding the column names.</param>
        /// <returns>A record batch with one column per schema entry and one row per data row.</returns>
        /// <exception cref="ArgumentNullException"><paramref name="sheet"/> is <see langword="null"/>.</exception>
        public static RecordBatch ToArrowRecordBatch(this IExcelSheet sheet, ExcelColumnSchema[]? schema = null, int headerRow = 1)
        {
            ArgumentNullException.ThrowIfNull(sheet);
            schema ??= Excel.InferSchema(sheet, headerRow);
            using IExcelRowEnumerator rows = sheet.GetEnumerator();
            return Build(rows, sheet.IsDate1904, schema, headerRow);
        }

        private static RecordBatch Build(IExcelRowEnumerator rows, bool isDate1904, ExcelColumnSchema[] resolvedSchema, int headerRow)
        {
            ColumnAppender[] appenders = new ColumnAppender[resolvedSchema.Length];
            for (int i = 0; i < resolvedSchema.Length; i++)
            {
                appenders[i] = ColumnAppender.Create(resolvedSchema[i]);
            }

            SkipHeaderRow(rows, headerRow);

            int rowCount = 0;
            while (rows.MoveNext())
            {
                Row row = rows.Current;
                for (int i = 0; i < appenders.Length; i++)
                {
                    appenders[i].Append(row[resolvedSchema[i].Index], isDate1904);
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
