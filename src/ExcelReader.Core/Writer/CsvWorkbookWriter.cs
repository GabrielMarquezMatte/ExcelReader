using System.Diagnostics.CodeAnalysis;

namespace ExcelReader.Core.Writer
{
    /// <summary>
    /// The single sheet of a CSV "workbook". Adapts the flat <see cref="CsvWriter"/> to the
    /// <see cref="ISheetWriter{TRow}"/> contract that <see cref="WorkbookRecordWriter{TSheet,TRow}"/> drives.
    /// </summary>
    /// <remarks>
    /// A CSV file is a single sheet, so the owning <see cref="CsvWorkbookWriter"/> exposes exactly one
    /// sheet and rejects a second <see cref="CsvWorkbookWriter.AddSheet"/> call. The workbook owns the
    /// <see cref="CsvWriter"/>; this sheet only borrows it.
    /// </remarks>
    public sealed class CsvSheetWriter : ISheetWriter<CsvRowWriter>
    {
        private readonly CsvWriter _writer;

        internal CsvSheetWriter(CsvWriter writer)
        {
            _writer = writer;
        }

        /// <inheritdoc/>
        public ValueTask StartAsync(CancellationToken ct = default)
        {
            return ValueTask.CompletedTask;
        }

        /// <inheritdoc/>
        public ValueTask<CsvRowWriter> StartRowAsync(CancellationToken ct = default)
        {
            ct.ThrowIfCancellationRequested();
            return new ValueTask<CsvRowWriter>(_writer.StartRow());
        }

        /// <inheritdoc/>
        public ValueTask EndAsync(CancellationToken ct = default)
        {
            return _writer.FlushAsync(ct);
        }

        /// <inheritdoc/>
        public ValueTask DisposeAsync()
        {
            // The workbook owns the CsvWriter's lifetime; nothing to release here.
            return ValueTask.CompletedTask;
        }
    }

    /// <summary>
    /// Writes a CSV file through the <see cref="IWorkbookWriter{TSheet}"/> contract. A CSV file holds
    /// exactly one sheet, so only a single <see cref="AddSheet"/> call is supported.
    /// </summary>
    public sealed class CsvWorkbookWriter : IWorkbookWriter<CsvSheetWriter>
    {
        private readonly CsvWriter _writer;
        private bool _sheetAdded;

        private CsvWorkbookWriter(CsvWriter writer)
        {
            _writer = writer;
        }

        /// <summary>
        /// Creates a writer that emits a CSV file to <paramref name="stream"/>.
        /// </summary>
        /// <param name="stream">The destination stream.</param>
        /// <param name="leaveOpen">If <see langword="true"/>, the stream is left open when the writer is disposed.</param>
        /// <param name="options">Delimiter/quote options; defaults to <see cref="CsvWriterOptions.Default"/> when omitted.</param>
        [SuppressMessage("Reliability", "CA2000:Dispose objects before losing scope",
            Justification = "Ownership of the CsvWriter transfers to CsvWorkbookWriter, which disposes it.")]
        [SuppressMessage("IDisposableAnalyzers.Correctness", "IDISP004:Don't ignore created IDisposable",
            Justification = "The created CsvWriter is stored in the returned workbook's field and disposed there.")]
        public static CsvWorkbookWriter Create(Stream stream, bool leaveOpen = false, CsvWriterOptions? options = null)
        {
            return new CsvWorkbookWriter(CsvWriter.Create(stream, leaveOpen, options));
        }

        /// <inheritdoc/>
        public ValueTask StartAsync(CancellationToken ct = default)
        {
            return ValueTask.CompletedTask;
        }

        /// <inheritdoc/>
        /// <exception cref="InvalidOperationException">A sheet was already added; a CSV file holds only one.</exception>
        public CsvSheetWriter AddSheet(string name)
        {
            if (_sheetAdded)
            {
                throw new InvalidOperationException("A CSV file holds a single sheet; only one AddSheet call is supported.");
            }
            _sheetAdded = true;
            return new CsvSheetWriter(_writer);
        }

        /// <inheritdoc/>
        public ValueTask EndAsync(CancellationToken ct = default)
        {
            return _writer.FlushAsync(ct);
        }

        /// <inheritdoc/>
        public ValueTask FlushAsync(CancellationToken ct = default)
        {
            return _writer.FlushAsync(ct);
        }

        /// <inheritdoc/>
        public ValueTask DisposeAsync()
        {
            return _writer.DisposeAsync();
        }
    }
}
