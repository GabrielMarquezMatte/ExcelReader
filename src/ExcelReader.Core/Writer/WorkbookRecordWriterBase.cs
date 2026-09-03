namespace ExcelReader.Core.Writer
{
    /// <summary>
    /// The sheet-lifecycle half of a record writer: owns the workbook writer, enforces unique sheet
    /// names, and drives start-sheet/header/rows/end-sheet. Derived types differ only in where a
    /// record type's headers and per-row write action come from — reflection
    /// (<see cref="WorkbookRecordWriter{TSheet,TRow}"/>) or <c>IExcelRecordMap{T}.ConfigureExcelRecordMap</c>
    /// (<see cref="MappedWorkbookRecordWriter{TSheet,TRow}"/>).
    /// </summary>
    /// <typeparam name="TSheet">The concrete sheet writer type.</typeparam>
    /// <typeparam name="TRow">The concrete row writer type.</typeparam>
    /// <remarks>
    /// Headers and the write action arrive as plain values, so nothing here needs
    /// <c>[RequiresUnreferencedCode]</c>/<c>[RequiresDynamicCode]</c> — only the reflection-backed
    /// derived type carries those. The constructor is <see langword="private protected"/>, so the two
    /// writers in this assembly are the only implementations; the type is public only because a public
    /// sealed class cannot derive from an internal one.
    /// </remarks>
    public abstract class WorkbookRecordWriterBase<TSheet, TRow> : IAsyncDisposable
        where TSheet : ISheetWriter<TRow>
        where TRow : IRowWriter
    {
        private readonly IWorkbookWriter<TSheet> _workbook;
        private readonly HashSet<string> _sheetNames = new(StringComparer.OrdinalIgnoreCase);

        private protected WorkbookRecordWriterBase(IWorkbookWriter<TSheet> workbook)
        {
            ArgumentNullException.ThrowIfNull(workbook);
            _workbook = workbook;
        }

        private protected async ValueTask WriteSheetCoreAsync<T>(string sheetName, IEnumerable<T> records,
                                                                 string[] headers, Action<TRow, T> writeRow,
                                                                 CancellationToken ct)
        {
            ArgumentNullException.ThrowIfNull(records);
            TSheet sheet = BeginSheet(sheetName);
            await using (sheet.ConfigureAwait(false))
            {
                await sheet.StartAsync(ct).ConfigureAwait(false);
                await WriteHeaderAsync(sheet, headers, ct).ConfigureAwait(false);
                await sheet.WriteRecordsAsync(records, writeRow, ct).ConfigureAwait(false);
                await sheet.EndAsync(ct).ConfigureAwait(false);
            }
        }

        private protected async ValueTask WriteSheetCoreAsync<T>(string sheetName, IAsyncEnumerable<T> records,
                                                                 string[] headers, Action<TRow, T> writeRow,
                                                                 CancellationToken ct)
        {
            ArgumentNullException.ThrowIfNull(records);
            TSheet sheet = BeginSheet(sheetName);
            await using (sheet.ConfigureAwait(false))
            {
                await sheet.StartAsync(ct).ConfigureAwait(false);
                await WriteHeaderAsync(sheet, headers, ct).ConfigureAwait(false);
                await sheet.WriteRecordsAsync(records, writeRow, ct).ConfigureAwait(false);
                await sheet.EndAsync(ct).ConfigureAwait(false);
            }
        }

        private TSheet BeginSheet(string sheetName)
        {
            ArgumentNullException.ThrowIfNull(sheetName);
            if (!_sheetNames.Add(sheetName))
            {
                throw new InvalidOperationException($"A sheet named '{sheetName}' already exists in this workbook.");
            }
            return _workbook.AddSheet(sheetName);
        }

        private static async ValueTask WriteHeaderAsync(TSheet sheet, string[] headers, CancellationToken ct)
        {
            TRow row = await sheet.StartRowAsync(ct).ConfigureAwait(false);
            await using (row.ConfigureAwait(false))
            {
                foreach (string header in headers)
                {
                    row.Write(header);
                }
            }
        }

        /// <summary>Finalizes and disposes the underlying workbook writer, completing the workbook.</summary>
        /// <remarks>Each workbook writer's DisposeAsync ends the workbook (finalizing all sheets) when started.</remarks>
        public async ValueTask DisposeAsync()
        {
            await DisposeAsyncCore().ConfigureAwait(false);
            GC.SuppressFinalize(this);
        }

        /// <summary>Releases the underlying workbook writer. Override to release a derived type's own resources first.</summary>
        protected virtual ValueTask DisposeAsyncCore()
        {
            return _workbook.DisposeAsync();
        }
    }
}
