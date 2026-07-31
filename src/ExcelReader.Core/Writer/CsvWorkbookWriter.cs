using System.Diagnostics.CodeAnalysis;
using ExcelReader.Core.Writer.Internal;

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
        private WriterState _state = WriterState.Created;
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
        [SuppressMessage("IDisposableAnalyzers.Correctness", "IDISP004:Don't ignore created IDisposable",
            Justification = "The created CsvWriter is stored in the returned workbook's field and disposed there.")]
        public static CsvWorkbookWriter Create(Stream stream, bool leaveOpen = false, CsvWriterOptions? options = null)
        {
            return new CsvWorkbookWriter(CsvWriter.Create(stream, leaveOpen, options));
        }

        /// <inheritdoc/>
        /// <exception cref="ObjectDisposedException">The workbook has already been ended.</exception>
        /// <exception cref="InvalidOperationException">The workbook has already been started.</exception>
        public ValueTask StartAsync(CancellationToken ct = default)
        {
            // COR-6: this used to do nothing at all — no state tracking, so AddSheet worked identically
            // before StartAsync, after EndAsync, or called twice, unlike the other three workbook
            // writers. WriterStateGuard is the shared machinery those three already use for exactly
            // this.
            WriterStateGuard.ThrowIfEnded(_state, this);
            WriterStateGuard.RequireCreated(_state, nameof(CsvWorkbookWriter));
            ct.ThrowIfCancellationRequested();
            _state = WriterState.Started;
            return ValueTask.CompletedTask;
        }

        /// <inheritdoc/>
        /// <exception cref="ArgumentNullException"><paramref name="name"/> is <see langword="null"/>.</exception>
        /// <exception cref="ArgumentException"><paramref name="name"/> is empty, longer than 31 characters, or contains one of <c>: \ / ? * [ ]</c>.</exception>
        /// <exception cref="ObjectDisposedException">The workbook has already been ended.</exception>
        /// <exception cref="InvalidOperationException">The workbook has not been started, or a sheet was already added; a CSV file holds only one.</exception>
        public CsvSheetWriter AddSheet(string name)
        {
            // sheetActive: false — CSV's "only one sheet" rule is the permanent _sheetAdded flag below,
            // not a "previous sheet still open" check like the multi-sheet writers use this parameter
            // for. Reused here for the null check, ThrowIfEnded/RequireStarted, and sheet-name validation
            // every other writer already gets from this same call.
            WriterStateGuard.RequireCanAddSheet(_state, this, nameof(CsvWorkbookWriter), name, sheetActive: false, nameof(CsvSheetWriter));
            if (_sheetAdded)
            {
                throw new InvalidOperationException("A CSV file holds a single sheet; only one AddSheet call is supported.");
            }
            _sheetAdded = true;
            return new CsvSheetWriter(_writer);
        }

        /// <inheritdoc/>
        /// <exception cref="ObjectDisposedException">The workbook has already been ended.</exception>
        /// <exception cref="InvalidOperationException">The workbook has not been started.</exception>
        public ValueTask EndAsync(CancellationToken ct = default)
        {
            WriterStateGuard.ThrowIfEnded(_state, this);
            WriterStateGuard.RequireStarted(_state, nameof(CsvWorkbookWriter), "ending");
            ct.ThrowIfCancellationRequested();
            _state = WriterState.Ended;
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
            _state = WriterState.Ended;
            return _writer.DisposeAsync();
        }
    }
}
