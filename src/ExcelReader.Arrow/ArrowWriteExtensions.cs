using System.Text;
using Apache.Arrow;
using Apache.Arrow.Types;
using ExcelReader.Core.Writer;

namespace ExcelReader.Arrow
{
    /// <summary>
    /// Writes an Apache Arrow <see cref="RecordBatch"/> to a workbook as one sheet — the write-side
    /// mirror of <see cref="ArrowConversionExtensions.ToArrowRecordBatch"/>.
    /// </summary>
    public static class ArrowWriteExtensions
    {
        /// <summary>Index 1 is always the builtin date style (see <see cref="IWorkbookWriter{TSheet}.AddStyle"/>).</summary>
        private const int BuiltinDateStyleId = 1;

        /// <summary>
        /// Writes <paramref name="batch"/> to <paramref name="workbook"/> as one XLSX sheet, in a single
        /// call: adds the sheet, writes the header row (unless <paramref name="writeHeader"/> is false)
        /// and every data row, then ends the sheet and the workbook.
        /// </summary>
        /// <param name="workbook">A freshly created, not-yet-started workbook writer.</param>
        /// <param name="batch">
        /// The batch to write. Only the seven Arrow types <see cref="ArrowConversionExtensions.ToArrowRecordBatch"/>
        /// produces (string, int64, double, boolean, date32, time64, timestamp) are supported.
        /// </param>
        /// <param name="sheetName">The sheet's name.</param>
        /// <param name="writeHeader">Whether to write a header row of column names first.</param>
        /// <exception cref="ArgumentNullException"><paramref name="workbook"/> or <paramref name="batch"/> is <see langword="null"/>.</exception>
        /// <exception cref="NotSupportedException">A column's Arrow type is not one of the seven supported types.</exception>
        public static void WriteRecordBatch(this XlsxWorkbookWriter workbook, RecordBatch batch, string sheetName = "Sheet1", bool writeHeader = true)
        {
            WriteRecordBatchCore<XlsxSheetWriter, XlsxRowWriter>(workbook, batch, sheetName, writeHeader);
        }

        /// <summary>Same as <see cref="WriteRecordBatch(XlsxWorkbookWriter, RecordBatch, string, bool)"/>, for an XLSB workbook.</summary>
        /// <exception cref="ArgumentNullException"><paramref name="workbook"/> or <paramref name="batch"/> is <see langword="null"/>.</exception>
        /// <exception cref="NotSupportedException">A column's Arrow type is not one of the seven supported types.</exception>
        public static void WriteRecordBatch(this XlsbWorkbookWriter workbook, RecordBatch batch, string sheetName = "Sheet1", bool writeHeader = true)
        {
            WriteRecordBatchCore<XlsbSheetWriter, XlsbRowWriter>(workbook, batch, sheetName, writeHeader);
        }

        /// <summary>Same as <see cref="WriteRecordBatch(XlsxWorkbookWriter, RecordBatch, string, bool)"/>, for a legacy XLS workbook.</summary>
        /// <exception cref="ArgumentNullException"><paramref name="workbook"/> or <paramref name="batch"/> is <see langword="null"/>.</exception>
        /// <exception cref="NotSupportedException">A column's Arrow type is not one of the seven supported types.</exception>
        public static void WriteRecordBatch(this XlsWorkbookWriter workbook, RecordBatch batch, string sheetName = "Sheet1", bool writeHeader = true)
        {
            WriteRecordBatchCore<XlsSheetWriter, XlsRowWriter>(workbook, batch, sheetName, writeHeader);
        }

        /// <summary>
        /// Writes <paramref name="batch"/> to <paramref name="workbook"/> as CSV's single unnamed sheet,
        /// in a single call.
        /// </summary>
        /// <exception cref="ArgumentNullException"><paramref name="workbook"/> or <paramref name="batch"/> is <see langword="null"/>.</exception>
        /// <exception cref="NotSupportedException">A column's Arrow type is not one of the seven supported types.</exception>
        public static void WriteRecordBatch(this CsvWorkbookWriter workbook, RecordBatch batch, bool writeHeader = true)
        {
            WriteRecordBatchCore<CsvSheetWriter, CsvRowWriter>(workbook, batch, sheetName: "Sheet1", writeHeader);
        }

        /// <summary>Asynchronous counterpart to <see cref="WriteRecordBatch(XlsxWorkbookWriter, RecordBatch, string, bool)"/>.</summary>
        public static ValueTask WriteRecordBatchAsync(this XlsxWorkbookWriter workbook, RecordBatch batch, string sheetName = "Sheet1", bool writeHeader = true, CancellationToken ct = default)
        {
            return WriteRecordBatchCoreAsync<XlsxSheetWriter, XlsxRowWriter>(workbook, batch, sheetName, writeHeader, ct);
        }

        /// <summary>Asynchronous counterpart to <see cref="WriteRecordBatch(XlsbWorkbookWriter, RecordBatch, string, bool)"/>.</summary>
        public static ValueTask WriteRecordBatchAsync(this XlsbWorkbookWriter workbook, RecordBatch batch, string sheetName = "Sheet1", bool writeHeader = true, CancellationToken ct = default)
        {
            return WriteRecordBatchCoreAsync<XlsbSheetWriter, XlsbRowWriter>(workbook, batch, sheetName, writeHeader, ct);
        }

        /// <summary>Asynchronous counterpart to <see cref="WriteRecordBatch(XlsWorkbookWriter, RecordBatch, string, bool)"/>.</summary>
        public static ValueTask WriteRecordBatchAsync(this XlsWorkbookWriter workbook, RecordBatch batch, string sheetName = "Sheet1", bool writeHeader = true, CancellationToken ct = default)
        {
            return WriteRecordBatchCoreAsync<XlsSheetWriter, XlsRowWriter>(workbook, batch, sheetName, writeHeader, ct);
        }

        /// <summary>Asynchronous counterpart to <see cref="WriteRecordBatch(CsvWorkbookWriter, RecordBatch, bool)"/>.</summary>
        public static ValueTask WriteRecordBatchAsync(this CsvWorkbookWriter workbook, RecordBatch batch, bool writeHeader = true, CancellationToken ct = default)
        {
            return WriteRecordBatchCoreAsync<CsvSheetWriter, CsvRowWriter>(workbook, batch, sheetName: "Sheet1", writeHeader, ct);
        }

        private static void WriteRecordBatchCore<TSheet, TRow>(IWorkbookWriter<TSheet> workbook, RecordBatch batch, string sheetName, bool writeHeader)
            where TSheet : ISheetWriter<TRow>
            where TRow : IRowWriter
        {
            ArgumentNullException.ThrowIfNull(workbook);
            ArgumentNullException.ThrowIfNull(batch);
            IReadOnlyList<Field> fields = batch.Schema.FieldsList;

            workbook.Start();
            TSheet sheet = workbook.AddSheet(sheetName);
            ApplyTemporalStyles<TSheet, TRow>(workbook, sheet, fields);
            sheet.Start();

            if (writeHeader)
            {
                using TRow header = sheet.StartRow();
                foreach (Field field in fields)
                {
                    header.Write(field.Name);
                }
            }
            for (int rowIndex = 0; rowIndex < batch.Length; rowIndex++)
            {
                using TRow row = sheet.StartRow();
                for (int col = 0; col < fields.Count; col++)
                {
                    WriteCell(row, batch.Column(col), fields[col].DataType.TypeId, rowIndex);
                }
            }

            sheet.End();
            sheet.Dispose();
            workbook.End();
        }

        private static async ValueTask WriteRecordBatchCoreAsync<TSheet, TRow>(IWorkbookWriter<TSheet> workbook, RecordBatch batch, string sheetName, bool writeHeader, CancellationToken ct)
            where TSheet : ISheetWriter<TRow>
            where TRow : IRowWriter
        {
            ArgumentNullException.ThrowIfNull(workbook);
            ArgumentNullException.ThrowIfNull(batch);
            IReadOnlyList<Field> fields = batch.Schema.FieldsList;

            await workbook.StartAsync(ct).ConfigureAwait(false);
            TSheet sheet = workbook.AddSheet(sheetName);
            ApplyTemporalStyles<TSheet, TRow>(workbook, sheet, fields);
            await sheet.StartAsync(ct).ConfigureAwait(false);

            if (writeHeader)
            {
                TRow header = await sheet.StartRowAsync(ct).ConfigureAwait(false);
                foreach (Field field in fields)
                {
                    header.Write(field.Name);
                }
                await header.DisposeAsync().ConfigureAwait(false);
            }
            for (int rowIndex = 0; rowIndex < batch.Length; rowIndex++)
            {
                TRow row = await sheet.StartRowAsync(ct).ConfigureAwait(false);
                for (int col = 0; col < fields.Count; col++)
                {
                    WriteCell(row, batch.Column(col), fields[col].DataType.TypeId, rowIndex);
                }
                await row.DisposeAsync().ConfigureAwait(false);
            }

            await sheet.EndAsync(ct).ConfigureAwait(false);
            await sheet.DisposeAsync().ConfigureAwait(false);
            await workbook.EndAsync(ct).ConfigureAwait(false);
        }

        // Time64/Timestamp cells write as plain numbers by default (see IRowWriter.Write(TimeOnly)'s
        // remarks) — without an explicit numFmt they render as a bare number in Excel. Date32 already
        // defaults to the builtin date style (IRowWriter.Write(DateOnly) falls back to style 1), so this
        // only reasserts it for clarity alongside the two columns that actually need the style call.
        private static void ApplyTemporalStyles<TSheet, TRow>(IWorkbookWriter<TSheet> workbook, TSheet sheet, IReadOnlyList<Field> fields)
            where TSheet : ISheetWriter<TRow>
            where TRow : IRowWriter
        {
            int timeStyle = -1;
            int timestampStyle = -1;
            for (int i = 0; i < fields.Count; i++)
            {
                switch (fields[i].DataType.TypeId)
                {
                    case ArrowTypeId.Date32:
                        sheet.SetColumnStyle(i, BuiltinDateStyleId);
                        break;
                    case ArrowTypeId.Time64:
                        timeStyle = timeStyle < 0 ? workbook.AddStyle(new CellStyle { NumberFormat = "hh:mm:ss" }) : timeStyle;
                        sheet.SetColumnStyle(i, timeStyle);
                        break;
                    case ArrowTypeId.Timestamp:
                        timestampStyle = timestampStyle < 0 ? workbook.AddStyle(new CellStyle { NumberFormat = "yyyy-mm-dd hh:mm:ss" }) : timestampStyle;
                        sheet.SetColumnStyle(i, timestampStyle);
                        break;
                    default:
                        break;
                }
            }
        }

        private static void WriteCell(IRowWriter row, IArrowArray array, ArrowTypeId typeId, int index)
        {
            switch (typeId)
            {
                case ArrowTypeId.String:
                    var strings = (StringArray)array;
                    row.Write(strings.IsNull(index) ? null : strings.GetString(index, Encoding.UTF8));
                    return;
                case ArrowTypeId.Int64:
                    row.Write(((Int64Array)array).GetValue(index));
                    return;
                case ArrowTypeId.Double:
                    row.Write(((DoubleArray)array).GetValue(index));
                    return;
                case ArrowTypeId.Boolean:
                    row.Write(((BooleanArray)array).GetValue(index));
                    return;
                case ArrowTypeId.Date32:
                    row.Write(((Date32Array)array).GetDateOnly(index));
                    return;
                case ArrowTypeId.Time64:
                    row.Write(((Time64Array)array).GetTime(index));
                    return;
                case ArrowTypeId.Timestamp:
                    // UtcDateTime, not DateTime: matches ArrowConversionExtensions' own read-side
                    // normalization (ToUniversalTime then strip Kind) for a batch built by another
                    // producer whose timestamps carry a real, non-zero offset.
                    row.Write(((TimestampArray)array).GetTimestamp(index)?.UtcDateTime);
                    return;
                default:
                    throw new NotSupportedException($"Arrow type {typeId} is not supported for writing.");
            }
        }
    }
}
