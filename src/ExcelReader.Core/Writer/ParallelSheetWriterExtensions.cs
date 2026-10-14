using ExcelReader.Core.Writer.Parallel;
using ExcelReader.Core.Writer.Xlsb;
using ExcelReader.Core.Writer.Xlsx;

namespace ExcelReader.Core.Writer
{
    /// <summary>
    /// Record writes that format and compress batches of rows on several threads and write them in
    /// order. Memory is bounded by <see cref="ParallelWriteOptions.DegreeOfParallelism"/> batches of
    /// <see cref="ParallelWriteOptions.BatchRows"/> rows.
    /// </summary>
    /// <remarks>
    /// The source is enumerated by one flow at a time, never concurrently. A <c>writeRow</c> delegate runs on pool
    /// threads, several at once, so whatever it captures must be safe to use concurrently; the layouts
    /// from <see cref="ExcelRecordLayout"/> are. Rows may be written through the normal API before
    /// and after a call, but not while one is running. A call that throws, a cancelled one included,
    /// ends the sheet and its workbook. With shared strings on, string indexes depend on thread
    /// timing, so two runs give the same content in files that differ byte for byte.
    /// </remarks>
    public static class ParallelSheetWriterExtensions
    {
        /// <summary>Writes a header row from <paramref name="layout"/>, then one row per record, in parallel batches.</summary>
        /// <typeparam name="T">The record type being written.</typeparam>
        /// <param name="sheet">The sheet to write rows to.</param>
        /// <param name="records">The records to write, one row each, in enumeration order.</param>
        /// <param name="layout">The columns to write, from <see cref="ExcelRecordLayout"/>.</param>
        /// <param name="options">Degree of parallelism and batch size. Defaults to <see cref="ParallelWriteOptions.Default"/>.</param>
        /// <param name="ct">A token to cancel the operation between batches.</param>
        /// <exception cref="ArgumentNullException"><paramref name="sheet"/>, <paramref name="records"/> or <paramref name="layout"/> is <see langword="null"/>.</exception>
        /// <exception cref="ArgumentOutOfRangeException"><paramref name="options"/> holds a negative degree or a batch size below 1.</exception>
        /// <exception cref="ObjectDisposedException">The sheet has already been ended.</exception>
        /// <exception cref="InvalidOperationException">A row is still active on the sheet.</exception>
        /// <exception cref="ExcelReader.Core.Reader.ExcelLimitExceededException">The records do not fit in the sheet's 1,048,576 rows.</exception>
        /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled. Unlike a serial write cancelled between rows, this ends the sheet and its workbook.</exception>
        public static ValueTask WriteRecordsParallelAsync<T>(this XlsxSheetWriter sheet, IEnumerable<T> records,
            ExcelRecordLayout<T> layout, ParallelWriteOptions? options = null, CancellationToken ct = default)
        {
            ArgumentNullException.ThrowIfNull(sheet);
            ArgumentNullException.ThrowIfNull(records);
            ArgumentNullException.ThrowIfNull(layout);
            return ParallelRecordWriter.WriteRecordsAsync<XlsxSheetWriter, XlsxRowWriter, T>(sheet, new RecordSource<T>(records), layout, options, ct);
        }

        /// <inheritdoc cref="WriteRecordsParallelAsync{T}(XlsxSheetWriter, IEnumerable{T}, ExcelRecordLayout{T}, ParallelWriteOptions?, CancellationToken)"/>
        public static ValueTask WriteRecordsParallelAsync<T>(this XlsxSheetWriter sheet, IAsyncEnumerable<T> records,
            ExcelRecordLayout<T> layout, ParallelWriteOptions? options = null, CancellationToken ct = default)
        {
            ArgumentNullException.ThrowIfNull(sheet);
            ArgumentNullException.ThrowIfNull(records);
            ArgumentNullException.ThrowIfNull(layout);
            return ParallelRecordWriter.WriteRecordsAsync<XlsxSheetWriter, XlsxRowWriter, T>(sheet, new RecordSource<T>(records), layout, options, ct);
        }

        /// <summary>Writes one row per record in parallel batches, calling <paramref name="writeRow"/> to populate each row. No header is written.</summary>
        /// <typeparam name="T">The record type being written.</typeparam>
        /// <param name="sheet">The sheet to write rows to.</param>
        /// <param name="records">The records to write, one row each, in enumeration order.</param>
        /// <param name="writeRow">Populates a row's cells from a single record. It runs on pool threads, several at once.</param>
        /// <param name="options">Degree of parallelism and batch size. Defaults to <see cref="ParallelWriteOptions.Default"/>.</param>
        /// <param name="ct">A token to cancel the operation between batches.</param>
        /// <exception cref="ArgumentNullException"><paramref name="sheet"/>, <paramref name="records"/> or <paramref name="writeRow"/> is <see langword="null"/>.</exception>
        /// <exception cref="ArgumentOutOfRangeException"><paramref name="options"/> holds a negative degree or a batch size below 1.</exception>
        /// <exception cref="ObjectDisposedException">The sheet has already been ended.</exception>
        /// <exception cref="InvalidOperationException">A row is still active on the sheet.</exception>
        /// <exception cref="ExcelReader.Core.Reader.ExcelLimitExceededException">The records do not fit in the sheet's 1,048,576 rows.</exception>
        /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled. Unlike a serial write cancelled between rows, this ends the sheet and its workbook.</exception>
        public static ValueTask WriteRecordsParallelAsync<T>(this XlsxSheetWriter sheet, IEnumerable<T> records,
            Action<XlsxRowWriter, T> writeRow, ParallelWriteOptions? options = null, CancellationToken ct = default)
        {
            ArgumentNullException.ThrowIfNull(sheet);
            ArgumentNullException.ThrowIfNull(records);
            ArgumentNullException.ThrowIfNull(writeRow);
            return ParallelRecordWriter.WriteRowsAsync(sheet, new RecordSource<T>(records), writeRow, options, ct);
        }

        /// <inheritdoc cref="WriteRecordsParallelAsync{T}(XlsxSheetWriter, IEnumerable{T}, Action{XlsxRowWriter, T}, ParallelWriteOptions?, CancellationToken)"/>
        public static ValueTask WriteRecordsParallelAsync<T>(this XlsxSheetWriter sheet, IAsyncEnumerable<T> records,
            Action<XlsxRowWriter, T> writeRow, ParallelWriteOptions? options = null, CancellationToken ct = default)
        {
            ArgumentNullException.ThrowIfNull(sheet);
            ArgumentNullException.ThrowIfNull(records);
            ArgumentNullException.ThrowIfNull(writeRow);
            return ParallelRecordWriter.WriteRowsAsync(sheet, new RecordSource<T>(records), writeRow, options, ct);
        }

        /// <summary>
        /// Writes <paramref name="records"/> as an Excel table named <paramref name="name"/>: a header row from
        /// <paramref name="layout"/>, one row per record in parallel batches, then closes the table.
        /// </summary>
        /// <typeparam name="T">The record type being written.</typeparam>
        /// <param name="sheet">The sheet to write to.</param>
        /// <param name="records">The records to write, one row each, in enumeration order.</param>
        /// <param name="name">The table name; see <see cref="ITableSheetWriter{TRow}.BeginTable"/>.</param>
        /// <param name="layout">The columns to write, from <see cref="ExcelRecordLayout"/>.</param>
        /// <param name="tableOptions">Position and style. Defaults to <see cref="ExcelTableOptions.Default"/>.</param>
        /// <param name="options">Degree of parallelism and batch size. Defaults to <see cref="ParallelWriteOptions.Default"/>.</param>
        /// <param name="ct">A token to cancel the operation between batches.</param>
        /// <exception cref="ArgumentNullException"><paramref name="sheet"/>, <paramref name="records"/>, <paramref name="name"/> or <paramref name="layout"/> is <see langword="null"/>.</exception>
        /// <exception cref="ArgumentOutOfRangeException"><paramref name="options"/> holds a negative degree or a batch size below 1.</exception>
        /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled. Unlike a serial write cancelled between rows, this ends the sheet and its workbook.</exception>
        public static ValueTask WriteTableParallelAsync<T>(this XlsxSheetWriter sheet, IEnumerable<T> records, string name,
            ExcelRecordLayout<T> layout, ExcelTableOptions? tableOptions = null, ParallelWriteOptions? options = null, CancellationToken ct = default)
        {
            ArgumentNullException.ThrowIfNull(sheet);
            ArgumentNullException.ThrowIfNull(records);
            ArgumentNullException.ThrowIfNull(layout);
            return ParallelRecordWriter.WriteTableAsync<XlsxSheetWriter, XlsxRowWriter, T>(sheet, new RecordSource<T>(records), name, layout, tableOptions, options, ct);
        }

        /// <inheritdoc cref="WriteTableParallelAsync{T}(XlsxSheetWriter, IEnumerable{T}, string, ExcelRecordLayout{T}, ExcelTableOptions?, ParallelWriteOptions?, CancellationToken)"/>
        public static ValueTask WriteTableParallelAsync<T>(this XlsxSheetWriter sheet, IAsyncEnumerable<T> records, string name,
            ExcelRecordLayout<T> layout, ExcelTableOptions? tableOptions = null, ParallelWriteOptions? options = null, CancellationToken ct = default)
        {
            ArgumentNullException.ThrowIfNull(sheet);
            ArgumentNullException.ThrowIfNull(records);
            ArgumentNullException.ThrowIfNull(layout);
            return ParallelRecordWriter.WriteTableAsync<XlsxSheetWriter, XlsxRowWriter, T>(sheet, new RecordSource<T>(records), name, layout, tableOptions, options, ct);
        }

        /// <inheritdoc cref="WriteRecordsParallelAsync{T}(XlsxSheetWriter, IEnumerable{T}, ExcelRecordLayout{T}, ParallelWriteOptions?, CancellationToken)"/>
        public static ValueTask WriteRecordsParallelAsync<T>(this XlsbSheetWriter sheet, IEnumerable<T> records,
            ExcelRecordLayout<T> layout, ParallelWriteOptions? options = null, CancellationToken ct = default)
        {
            ArgumentNullException.ThrowIfNull(sheet);
            ArgumentNullException.ThrowIfNull(records);
            ArgumentNullException.ThrowIfNull(layout);
            return ParallelRecordWriter.WriteRecordsAsync<XlsbSheetWriter, XlsbRowWriter, T>(sheet, new RecordSource<T>(records), layout, options, ct);
        }

        /// <inheritdoc cref="WriteRecordsParallelAsync{T}(XlsxSheetWriter, IEnumerable{T}, ExcelRecordLayout{T}, ParallelWriteOptions?, CancellationToken)"/>
        public static ValueTask WriteRecordsParallelAsync<T>(this XlsbSheetWriter sheet, IAsyncEnumerable<T> records,
            ExcelRecordLayout<T> layout, ParallelWriteOptions? options = null, CancellationToken ct = default)
        {
            ArgumentNullException.ThrowIfNull(sheet);
            ArgumentNullException.ThrowIfNull(records);
            ArgumentNullException.ThrowIfNull(layout);
            return ParallelRecordWriter.WriteRecordsAsync<XlsbSheetWriter, XlsbRowWriter, T>(sheet, new RecordSource<T>(records), layout, options, ct);
        }

        /// <inheritdoc cref="WriteRecordsParallelAsync{T}(XlsxSheetWriter, IEnumerable{T}, Action{XlsxRowWriter, T}, ParallelWriteOptions?, CancellationToken)"/>
        public static ValueTask WriteRecordsParallelAsync<T>(this XlsbSheetWriter sheet, IEnumerable<T> records,
            Action<XlsbRowWriter, T> writeRow, ParallelWriteOptions? options = null, CancellationToken ct = default)
        {
            ArgumentNullException.ThrowIfNull(sheet);
            ArgumentNullException.ThrowIfNull(records);
            ArgumentNullException.ThrowIfNull(writeRow);
            return ParallelRecordWriter.WriteRowsAsync(sheet, new RecordSource<T>(records), writeRow, options, ct);
        }

        /// <inheritdoc cref="WriteRecordsParallelAsync{T}(XlsxSheetWriter, IEnumerable{T}, Action{XlsxRowWriter, T}, ParallelWriteOptions?, CancellationToken)"/>
        public static ValueTask WriteRecordsParallelAsync<T>(this XlsbSheetWriter sheet, IAsyncEnumerable<T> records,
            Action<XlsbRowWriter, T> writeRow, ParallelWriteOptions? options = null, CancellationToken ct = default)
        {
            ArgumentNullException.ThrowIfNull(sheet);
            ArgumentNullException.ThrowIfNull(records);
            ArgumentNullException.ThrowIfNull(writeRow);
            return ParallelRecordWriter.WriteRowsAsync(sheet, new RecordSource<T>(records), writeRow, options, ct);
        }

        /// <inheritdoc cref="WriteTableParallelAsync{T}(XlsxSheetWriter, IEnumerable{T}, string, ExcelRecordLayout{T}, ExcelTableOptions?, ParallelWriteOptions?, CancellationToken)"/>
        public static ValueTask WriteTableParallelAsync<T>(this XlsbSheetWriter sheet, IEnumerable<T> records, string name,
            ExcelRecordLayout<T> layout, ExcelTableOptions? tableOptions = null, ParallelWriteOptions? options = null, CancellationToken ct = default)
        {
            ArgumentNullException.ThrowIfNull(sheet);
            ArgumentNullException.ThrowIfNull(records);
            ArgumentNullException.ThrowIfNull(layout);
            return ParallelRecordWriter.WriteTableAsync<XlsbSheetWriter, XlsbRowWriter, T>(sheet, new RecordSource<T>(records), name, layout, tableOptions, options, ct);
        }

        /// <inheritdoc cref="WriteTableParallelAsync{T}(XlsxSheetWriter, IEnumerable{T}, string, ExcelRecordLayout{T}, ExcelTableOptions?, ParallelWriteOptions?, CancellationToken)"/>
        public static ValueTask WriteTableParallelAsync<T>(this XlsbSheetWriter sheet, IAsyncEnumerable<T> records, string name,
            ExcelRecordLayout<T> layout, ExcelTableOptions? tableOptions = null, ParallelWriteOptions? options = null, CancellationToken ct = default)
        {
            ArgumentNullException.ThrowIfNull(sheet);
            ArgumentNullException.ThrowIfNull(records);
            ArgumentNullException.ThrowIfNull(layout);
            return ParallelRecordWriter.WriteTableAsync<XlsbSheetWriter, XlsbRowWriter, T>(sheet, new RecordSource<T>(records), name, layout, tableOptions, options, ct);
        }
    }
}
