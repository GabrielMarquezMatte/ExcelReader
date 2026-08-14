using System.Diagnostics.CodeAnalysis;
using System.IO.Compression;

namespace ExcelReader.Core.Writer
{
    /// <summary>
    /// Writes plain-old-CLR-object records to a workbook as sheets, exactly like
    /// <see cref="WorkbookRecordWriter{TSheet,TRow}"/>, but from a map <c>IExcelRecordMap{T}.ConfigureExcelRecordMap</c>
    /// builds (source-generated or hand-written) instead of reflecting over the record type.
    /// </summary>
    /// <typeparam name="TSheet">The concrete sheet writer type.</typeparam>
    /// <typeparam name="TRow">The concrete row writer type.</typeparam>
    /// <remarks>
    /// No <c>[RequiresUnreferencedCode]</c>/<c>[RequiresDynamicCode]</c>: the <c>where T : IExcelRecordMap&lt;T&gt;</c>
    /// constraint on <see cref="WriteSheetAsync{T}(string, IEnumerable{T}, CancellationToken)"/> guarantees
    /// the column plan comes from the record type's own <c>IExcelRecordMap{T}.ConfigureExcelRecordMap</c>,
    /// which reaches neither <c>GetProperties</c> nor <c>Expression.Compile</c>/<c>MakeGenericMethod</c>.
    /// </remarks>
    public sealed class MappedWorkbookRecordWriter<TSheet, TRow> : IAsyncDisposable where TSheet : ISheetWriter<TRow> where TRow : IRowWriter
    {
        private readonly IWorkbookWriter<TSheet> _workbook;
        private readonly HashSet<string> _sheetNames = new(StringComparer.OrdinalIgnoreCase);

        /// <summary>
        /// Wraps an already-created, already-started workbook writer. Ownership of <paramref name="workbook"/>
        /// transfers to this instance, which disposes it when this instance is disposed.
        /// </summary>
        /// <param name="workbook">The started workbook writer to wrap.</param>
        public MappedWorkbookRecordWriter(IWorkbookWriter<TSheet> workbook)
        {
            ArgumentNullException.ThrowIfNull(workbook);
            _workbook = workbook;
        }

        /// <summary>
        /// Writes a new sheet named <paramref name="sheetName"/> containing a header row followed by one
        /// row per item in <paramref name="records"/>.
        /// </summary>
        /// <typeparam name="T">The record type; must implement <see cref="IExcelRecordMap{T}"/>.</typeparam>
        /// <param name="sheetName">The sheet's name; must be unique within this workbook.</param>
        /// <param name="records">The records to write, one row each, in enumeration order.</param>
        /// <param name="ct">A token to cancel the operation.</param>
        /// <exception cref="InvalidOperationException">A sheet named <paramref name="sheetName"/> already exists in this workbook.</exception>
        public async ValueTask WriteSheetAsync<T>(string sheetName, IEnumerable<T> records, CancellationToken ct = default)
            where T : IExcelRecordMap<T>
        {
            ArgumentNullException.ThrowIfNull(records);
            TSheet sheet = BeginSheet(sheetName);
            await using (sheet.ConfigureAwait(false))
            {
                await sheet.StartAsync(ct).ConfigureAwait(false);
                await WriteHeaderAsync<T>(sheet, ct).ConfigureAwait(false);
                await sheet.WriteRecordsAsync(records, static (row, record) => MappedRecordColumns<T>.WriteRow(row, record), ct).ConfigureAwait(false);
                await sheet.EndAsync(ct).ConfigureAwait(false);
            }
        }

        /// <summary>
        /// Writes a new sheet named <paramref name="sheetName"/> containing a header row followed by one
        /// row per item produced by <paramref name="records"/>.
        /// </summary>
        /// <typeparam name="T">The record type; must implement <see cref="IExcelRecordMap{T}"/>.</typeparam>
        /// <param name="sheetName">The sheet's name; must be unique within this workbook.</param>
        /// <param name="records">The records to write, one row each, in enumeration order.</param>
        /// <param name="ct">A token to cancel the operation, and passed to the source enumerable.</param>
        /// <exception cref="InvalidOperationException">A sheet named <paramref name="sheetName"/> already exists in this workbook.</exception>
        public async ValueTask WriteSheetAsync<T>(string sheetName, IAsyncEnumerable<T> records, CancellationToken ct = default)
            where T : IExcelRecordMap<T>
        {
            ArgumentNullException.ThrowIfNull(records);
            TSheet sheet = BeginSheet(sheetName);
            await using (sheet.ConfigureAwait(false))
            {
                await sheet.StartAsync(ct).ConfigureAwait(false);
                await WriteHeaderAsync<T>(sheet, ct).ConfigureAwait(false);
                await sheet.WriteRecordsAsync(records, static (row, record) => MappedRecordColumns<T>.WriteRow(row, record), ct).ConfigureAwait(false);
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

        private static async ValueTask WriteHeaderAsync<T>(TSheet sheet, CancellationToken ct) where T : IExcelRecordMap<T>
        {
            TRow row = await sheet.StartRowAsync(ct).ConfigureAwait(false);
            await using (row.ConfigureAwait(false))
            {
                foreach (string header in MappedRecordColumns<T>.Headers)
                {
                    row.Write(header);
                }
            }
        }

        /// <summary>Finalizes and disposes the underlying workbook writer, completing the workbook.</summary>
        [SuppressMessage("IDisposableAnalyzers.Correctness", "IDISP007:Don't dispose injected",
            Justification = "The RecordWriter factories transfer ownership of the workbook to this wrapper, which finalizes it here.")]
        public ValueTask DisposeAsync()
        {
            return _workbook.DisposeAsync();
        }
    }

    // The mapped-path counterpart to RecordColumns<T>: instead of reflecting over T's properties, builds
    // T's ExcelRecordMapBuilder<T> once (via T.ConfigureExcelRecordMap, source-generated or hand-written)
    // and caches it for the process lifetime. Headers/WriteRow don't depend on the concrete row-writer
    // type the way RecordColumns<T>'s Expression-compiled Plan<TRow> does, since ExcelRecordMapBuilder<T>
    // already compiles its column actions against the IRowWriter interface at the call site (ordinary C#
    // interface dispatch, not per-TRow Expression trees) — one plan per T covers every format.
    [SuppressMessage("Major Code Smell", "S2743:Static fields should not be used in generic types",
        Justification = "The per-closed-type static IS the design: the map is built once per T, not shared across different T.")]
    internal static class MappedRecordColumns<T> where T : IExcelRecordMap<T>
    {
        private static readonly ExcelRecordMapBuilder<T> _builder = Build();

        internal static string[] Headers { get; } = _builder.Headers();

        internal static void WriteRow(IRowWriter row, T record)
        {
            _builder.WriteRow(row, record);
        }

        private static ExcelRecordMapBuilder<T> Build()
        {
            var builder = new ExcelRecordMapBuilder<T>();
            T.ConfigureExcelRecordMap(builder);
            return builder;
        }
    }

    /// <summary>
    /// Format-specific factories that create the underlying low-level workbook writer, start it, and wrap
    /// it in a <see cref="MappedWorkbookRecordWriter{TSheet,TRow}"/> — the AOT-clean counterpart to
    /// <see cref="RecordWriter"/>'s factories.
    /// </summary>
    public static class MappedRecordWriter
    {
        /// <summary>Creates and starts a record writer that produces an XLSX workbook.</summary>
        /// <param name="stream">The destination stream; must support writing.</param>
        /// <param name="leaveOpen">If <see langword="true"/>, <paramref name="stream"/> is left open when the returned writer is disposed.</param>
        /// <param name="compression">The ZIP compression level used for the XLSX package's entries.</param>
        /// <param name="useSharedStrings">If <see langword="true"/>, text cells are deduplicated into the shared-strings table instead of written inline.</param>
        /// <param name="prefetchWrite">If <see langword="true"/>, each sheet's deflate runs on a background thread instead of the calling thread (see <see cref="XlsxWorkbookWriter.CreateAsync"/>).</param>
        /// <param name="ct">A token to cancel the operation.</param>
        /// <returns>A started record writer ready to accept sheets.</returns>
        public static async ValueTask<MappedWorkbookRecordWriter<XlsxSheetWriter, XlsxRowWriter>> CreateMappedXlsxAsync(
            Stream stream, bool leaveOpen = false, CompressionLevel compression = CompressionLevel.Fastest,
            bool useSharedStrings = false, bool prefetchWrite = false, CancellationToken ct = default)
        {
            var workbook = await XlsxWorkbookWriter.CreateAsync(stream, leaveOpen, compression, useSharedStrings, prefetchWrite, ct).ConfigureAwait(false);
            await workbook.StartAsync(ct).ConfigureAwait(false);
            return new MappedWorkbookRecordWriter<XlsxSheetWriter, XlsxRowWriter>(workbook);
        }

        /// <summary>Creates and starts a record writer that produces an XLSB workbook.</summary>
        /// <param name="stream">The destination stream; must support writing.</param>
        /// <param name="leaveOpen">If <see langword="true"/>, <paramref name="stream"/> is left open when the returned writer is disposed.</param>
        /// <param name="date1904">If <see langword="true"/>, dates are serialized using the 1904 date system instead of the default 1900 system.</param>
        /// <param name="compression">The ZIP compression level used for the XLSB package's entries.</param>
        /// <param name="useSharedStrings">If <see langword="true"/>, text cells are deduplicated into the shared-strings table instead of written inline.</param>
        /// <param name="prefetchWrite">If <see langword="true"/>, each sheet's deflate runs on a background thread instead of the calling thread (see <see cref="XlsbWorkbookWriter.CreateAsync"/>).</param>
        /// <param name="ct">A token to cancel the operation.</param>
        /// <returns>A started record writer ready to accept sheets.</returns>
        public static async ValueTask<MappedWorkbookRecordWriter<XlsbSheetWriter, XlsbRowWriter>> CreateMappedXlsbAsync(
            Stream stream, bool leaveOpen = false, bool date1904 = false,
            CompressionLevel compression = CompressionLevel.Fastest, bool useSharedStrings = false,
            bool prefetchWrite = false, CancellationToken ct = default)
        {
            var workbook = await XlsbWorkbookWriter.CreateAsync(stream, leaveOpen, date1904, compression, useSharedStrings, prefetchWrite, ct).ConfigureAwait(false);
            await workbook.StartAsync(ct).ConfigureAwait(false);
            return new MappedWorkbookRecordWriter<XlsbSheetWriter, XlsbRowWriter>(workbook);
        }

        /// <summary>
        /// Creates and starts a record writer that produces a CSV file. Supports only a single sheet,
        /// since a CSV file is inherently one sheet. The returned writer is still <see cref="IAsyncDisposable"/>.
        /// </summary>
        /// <param name="stream">The destination stream; must support writing.</param>
        /// <param name="leaveOpen">If <see langword="true"/>, <paramref name="stream"/> is left open when the returned writer is disposed.</param>
        /// <param name="options">The delimiter/quote character to use; defaults to <see cref="CsvWriterOptions.Default"/> if <see langword="null"/>.</param>
        /// <param name="ct">A token to cancel the operation.</param>
        /// <returns>A started record writer ready to accept its single sheet.</returns>
        [SuppressMessage("Reliability", "CA2000:Dispose objects before losing scope",
            Justification = "Ownership of the workbook transfers to MappedWorkbookRecordWriter, which disposes it.")]
        public static async ValueTask<MappedWorkbookRecordWriter<CsvSheetWriter, CsvRowWriter>> CreateMappedCsvAsync(
            Stream stream, bool leaveOpen = false, CsvWriterOptions? options = null, CancellationToken ct = default)
        {
            var workbook = CsvWorkbookWriter.Create(stream, leaveOpen, options);
            await workbook.StartAsync(ct).ConfigureAwait(false);
            return new MappedWorkbookRecordWriter<CsvSheetWriter, CsvRowWriter>(workbook);
        }

        /// <summary>Creates and starts a record writer that produces a legacy XLS (BIFF8) workbook.</summary>
        /// <param name="stream">The destination stream; must support writing.</param>
        /// <param name="leaveOpen">If <see langword="true"/>, <paramref name="stream"/> is left open when the returned writer is disposed.</param>
        /// <param name="date1904">If <see langword="true"/>, dates are serialized using the 1904 date system instead of the default 1900 system.</param>
        /// <param name="ct">A token to cancel the operation.</param>
        /// <returns>A started record writer ready to accept sheets.</returns>
        [SuppressMessage("Reliability", "CA2000:Dispose objects before losing scope",
            Justification = "Ownership of the workbook transfers to MappedWorkbookRecordWriter, which disposes it.")]
        public static async ValueTask<MappedWorkbookRecordWriter<XlsSheetWriter, XlsRowWriter>> CreateMappedXlsAsync(
            Stream stream, bool leaveOpen = false, bool date1904 = false, CancellationToken ct = default)
        {
            var workbook = XlsWorkbookWriter.Create(stream, leaveOpen, date1904);
            await workbook.StartAsync(ct).ConfigureAwait(false);
            return new MappedWorkbookRecordWriter<XlsSheetWriter, XlsRowWriter>(workbook);
        }
    }
}
