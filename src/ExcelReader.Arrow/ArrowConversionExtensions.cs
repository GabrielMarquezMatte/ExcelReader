using Apache.Arrow;
using ExcelReader.Core.Reader;
using ExcelReader.Core.Reader.Schema;

namespace ExcelReader.Arrow
{
    /// <summary>Converts an <see cref="IExcelSheet"/> into an Apache Arrow <see cref="RecordBatch"/>.</summary>
    public static class ArrowConversionExtensions
    {
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
            bool isDate1904 = sheet.IsDate1904;
            using IExcelRowEnumerator rows = sheet.GetEnumerator();
            ColumnAppender[] appenders = new ColumnAppender[schema.Length];
            for (int i = 0; i < schema.Length; i++)
            {
                appenders[i] = ColumnAppender.Create(schema[i]);
            }

            SkipHeaderRow(rows, headerRow);

            int rowCount = 0;
            while (rows.MoveNext())
            {
                Row row = rows.Current;
                for (int i = 0; i < appenders.Length; i++)
                {
                    appenders[i].Append(row[schema[i].Index], isDate1904);
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
