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

        /// <summary>
        /// Synchronous counterpart to <see cref="StartAsync"/>, for native/unmanaged callers whose ABI
        /// is synchronous. CSV has no leading structure to write, so this is a no-op.
        /// </summary>
        public void Start()
        {
        }

        /// <inheritdoc/>
        public ValueTask StartAsync(CancellationToken ct = default)
        {
            return ValueTask.CompletedTask;
        }

        /// <summary>
        /// Synchronous counterpart to <see cref="StartRowAsync(CancellationToken)"/>, for native/unmanaged
        /// callers whose ABI is synchronous.
        /// </summary>
        public CsvRowWriter StartRow()
        {
            return _writer.StartRow();
        }

        /// <summary>
        /// Validates <paramref name="styleId"/> is not negative, then no-ops on it: CSV has no cell
        /// styles. Synchronous counterpart to <see cref="StartRowAsync(int, CancellationToken)"/>, for
        /// native/unmanaged callers whose ABI is synchronous.
        /// </summary>
        /// <exception cref="ArgumentOutOfRangeException"><paramref name="styleId"/> is negative.</exception>
        public CsvRowWriter StartRow(int styleId)
        {
            ArgumentOutOfRangeException.ThrowIfNegative(styleId);
            return StartRow();
        }

        /// <inheritdoc/>
        public ValueTask<CsvRowWriter> StartRowAsync(CancellationToken ct = default)
        {
            ct.ThrowIfCancellationRequested();
            return new ValueTask<CsvRowWriter>(_writer.StartRow());
        }

        /// <summary>Validates <paramref name="styleId"/> is not negative, then no-ops: CSV has no cell styles.</summary>
        /// <inheritdoc cref="ISheetWriter{TRow}.StartRowAsync(int, CancellationToken)"/>
        /// <exception cref="ArgumentOutOfRangeException"><paramref name="styleId"/> is negative.</exception>
        public ValueTask<CsvRowWriter> StartRowAsync(int styleId, CancellationToken ct = default)
        {
            ArgumentOutOfRangeException.ThrowIfNegative(styleId);
            return StartRowAsync(ct);
        }

        /// <summary>Validates its arguments are not negative, then no-ops: CSV has no column styles.</summary>
        /// <inheritdoc cref="ISheetWriter{TRow}.SetColumnStyle"/>
        /// <exception cref="ArgumentOutOfRangeException"><paramref name="columnIndex"/> or <paramref name="styleId"/> is negative.</exception>
        public void SetColumnStyle(int columnIndex, int styleId)
        {
            ArgumentOutOfRangeException.ThrowIfNegative(columnIndex);
            ArgumentOutOfRangeException.ThrowIfNegative(styleId);
        }

        /// <summary>Validates its arguments, then no-ops: CSV has no column widths.</summary>
        /// <inheritdoc cref="ISheetWriter{TRow}.SetColumnWidth"/>
        /// <exception cref="ArgumentOutOfRangeException"><paramref name="columnIndex"/> or <paramref name="width"/> is negative.</exception>
        public void SetColumnWidth(int columnIndex, double width)
        {
            ArgumentOutOfRangeException.ThrowIfNegative(columnIndex);
            ArgumentOutOfRangeException.ThrowIfNegative(width);
        }

        /// <summary>
        /// Synchronous counterpart to <see cref="EndAsync"/>, for native/unmanaged callers whose ABI is
        /// synchronous.
        /// </summary>
        public void End()
        {
            _writer.Flush();
        }

        /// <inheritdoc/>
        public ValueTask EndAsync(CancellationToken ct = default)
        {
            return _writer.FlushAsync(ct);
        }

        /// <summary>
        /// Synchronous counterpart to <see cref="DisposeAsync"/>, for native/unmanaged callers whose ABI
        /// is synchronous. The workbook owns the <see cref="CsvWriter"/>'s lifetime; nothing to release
        /// here.
        /// </summary>
        public void Dispose()
        {
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
        public static CsvWorkbookWriter Create(Stream stream, bool leaveOpen = false, CsvWriterOptions? options = null)
        {
            return new CsvWorkbookWriter(CsvWriter.Create(stream, leaveOpen, options));
        }

        /// <summary>
        /// Synchronous counterpart to <see cref="StartAsync"/>, for native/unmanaged callers whose ABI
        /// is synchronous.
        /// </summary>
        /// <exception cref="ObjectDisposedException">The workbook has already been ended.</exception>
        /// <exception cref="InvalidOperationException">The workbook has already been started.</exception>
        public void Start()
        {
            WriterStateGuard.ThrowIfEnded(_state, this);
            WriterStateGuard.RequireCreated(_state, nameof(CsvWorkbookWriter));
            _state = WriterState.Started;
        }

        /// <inheritdoc/>
        /// <exception cref="ObjectDisposedException">The workbook has already been ended.</exception>
        /// <exception cref="InvalidOperationException">The workbook has already been started.</exception>
        public ValueTask StartAsync(CancellationToken ct = default)
        {
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
            // sheetActive: false — the "only one sheet" rule is the permanent _sheetAdded flag below.
            WriterStateGuard.RequireCanAddSheet(_state, this, nameof(CsvWorkbookWriter), name, sheetActive: false, nameof(CsvSheetWriter));
            if (_sheetAdded)
            {
                throw new InvalidOperationException("A CSV file holds a single sheet; only one AddSheet call is supported.");
            }
            _sheetAdded = true;
            return new CsvSheetWriter(_writer);
        }

        /// <summary>
        /// Synchronous counterpart to <see cref="EndAsync"/>, for native/unmanaged callers whose ABI is
        /// synchronous.
        /// </summary>
        /// <exception cref="ObjectDisposedException">The workbook has already been ended.</exception>
        /// <exception cref="InvalidOperationException">The workbook has not been started.</exception>
        public void End()
        {
            WriterStateGuard.ThrowIfEnded(_state, this);
            WriterStateGuard.RequireStarted(_state, nameof(CsvWorkbookWriter), "ending");
            _state = WriterState.Ended;
            _writer.Flush();
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

        /// <summary>
        /// Synchronous counterpart to <see cref="FlushAsync"/>, for native/unmanaged callers whose ABI
        /// is synchronous.
        /// </summary>
        public void Flush()
        {
            _writer.Flush();
        }

        /// <inheritdoc/>
        public ValueTask FlushAsync(CancellationToken ct = default)
        {
            return _writer.FlushAsync(ct);
        }

        /// <summary>No-op: CSV has no cell styles. Always returns 0.</summary>
        /// <inheritdoc cref="IWorkbookWriter{TSheet}.AddStyle"/>
        public int AddStyle(CellStyle style)
        {
            return 0;
        }

        /// <summary>
        /// Synchronous counterpart to <see cref="DisposeAsync"/>, for native/unmanaged callers whose ABI
        /// is synchronous.
        /// </summary>
        public void Dispose()
        {
            _state = WriterState.Ended;
            _writer.Dispose();
        }

        /// <inheritdoc/>
        public ValueTask DisposeAsync()
        {
            _state = WriterState.Ended;
            return _writer.DisposeAsync();
        }
    }
}
