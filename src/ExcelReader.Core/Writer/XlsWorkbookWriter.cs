using System.Runtime.InteropServices;
using ExcelReader.Core.Writer.Internal;

namespace ExcelReader.Core.Writer
{
    /// <summary>
    /// Writes a workbook to the legacy BIFF8 (.xls) binary format.
    /// </summary>
    /// <remarks>
    /// Structured the same way as <see cref="XlsxWorkbookWriter"/> (XLSX). Unlike the ZIP writer, records
    /// are buffered in memory and the OLE container is assembled in <see cref="EndAsync"/>: the BoundSheet
    /// offsets in the globals and the OLE FAT both need stream sizes known only at the end. Rows beyond
    /// the BIFF8 per-sheet cap (65,536) overflow into auto-generated continuation sheets, so memory scales
    /// with total row count rather than being bounded per sheet.
    /// </remarks>
    public sealed class XlsWorkbookWriter : IWorkbookWriter<XlsSheetWriter>
    {
        private readonly Stream _stream;
        private readonly bool _leaveOpen;
        private readonly bool _date1904;
        private readonly StyleTable _styles = new();
        private readonly List<XlsSheetWriter> _sheets = [];
        private WriterState _state = WriterState.Created;
        private XlsSheetWriter? _activeSheet;
        private bool _disposed;

        private XlsWorkbookWriter(Stream stream, bool leaveOpen, bool date1904)
        {
            _stream = stream;
            _leaveOpen = leaveOpen;
            _date1904 = date1904;
        }

        /// <summary>
        /// Creates a writer that emits a BIFF8 (.xls) workbook to <paramref name="stream"/>.
        /// </summary>
        /// <param name="stream">The destination stream.</param>
        /// <param name="leaveOpen">If <see langword="true"/>, the stream is left open when the writer is disposed.</param>
        /// <param name="date1904">If <see langword="true"/>, dates are serialized using the 1904 epoch instead of the default 1900 epoch.</param>
        public static XlsWorkbookWriter Create(Stream stream, bool leaveOpen = false, bool date1904 = false)
        {
            ArgumentNullException.ThrowIfNull(stream);
            return new XlsWorkbookWriter(stream, leaveOpen, date1904);
        }

        /// <summary>
        /// Marks the workbook as started so that sheets can be added.
        /// </summary>
        public void Start()
        {
            WriterStateGuard.ThrowIfEnded(_state, this);
            WriterStateGuard.RequireCreated(_state, nameof(XlsWorkbookWriter));
            _state = WriterState.Started;
        }

        /// <inheritdoc/>
        /// <remarks>XLS assembles the OLE container synchronously in <see cref="EndAsync"/>; this just wraps <see cref="Start"/>.</remarks>
        public ValueTask StartAsync(CancellationToken ct = default)
        {
            ct.ThrowIfCancellationRequested();
            Start();
            return ValueTask.CompletedTask;
        }

        /// <inheritdoc/>
        /// <exception cref="InvalidOperationException">The previously added sheet has not been ended yet.</exception>
        public XlsSheetWriter AddSheet(string name)
        {
            WriterStateGuard.RequireCanAddSheet(
                _state, this, nameof(XlsWorkbookWriter), name, _activeSheet is not null, nameof(XlsSheetWriter));
#pragma warning disable IDISP003 // The guard above guarantees _activeSheet is null here — there is no previous to dispose.
            _activeSheet = new XlsSheetWriter(this, name, _date1904);
#pragma warning restore IDISP003
            return _activeSheet;
        }

        internal int SheetCount => _sheets.Count;

        /// <inheritdoc/>
        public int AddStyle(CellStyle style)
        {
            return _styles.Add(style);
        }

        internal void RegisterSheet(XlsSheetWriter sheet)
        {
            _sheets.Add(sheet);
        }

        internal void NotifySheetEnded()
        {
            _activeSheet?.Dispose();
            _activeSheet = null;
        }

        /// <inheritdoc/>
        /// <exception cref="InvalidOperationException">No sheet was ever added to the workbook.</exception>
        public async ValueTask EndAsync(CancellationToken ct = default)
        {
            WriterStateGuard.ThrowIfEnded(_state, this);
            WriterStateGuard.RequireStarted(_state, nameof(XlsWorkbookWriter), "ending");
            ct.ThrowIfCancellationRequested();
            _state = WriterState.Ended;
            if (_activeSheet is not null)
            {
                await _activeSheet.DisposeAsync().ConfigureAwait(false);
            }
            if (_sheets.Count == 0)
            {
                throw new InvalidOperationException("A workbook must contain at least one sheet.");
            }

            // Globals are small (BOF/CODEPAGE/XFs/BoundSheets/EOF); only this buffer plus each
            // sheet's already-accumulated cell buffer live at finalize — no combined workbook copy.
            BiffBuffer globals = new(1024);
            BiffBuffer frame = new(64);
            try
            {
                string[] names = new string[_sheets.Count];
                for (int i = 0; i < _sheets.Count; i++)
                {
                    names[i] = _sheets[i].Name;
                }

                int[] offsetPositions = XlsGlobals.Write(globals, names, _date1904, _styles);
                int offset = globals.Length;
                int workbookSize = offset;
                for (int i = 0; i < _sheets.Count; i++)
                {
                    globals.PatchI32(offsetPositions[i], offset);
                    offset += _sheets[i].SubstreamLength;
                    workbookSize += _sheets[i].SubstreamLength;
                }

                await OleCompoundWriter.WriteAsync(_stream, workbookSize, async (dest, canc) =>
                {
                    await dest.WriteAsync(globals.Memory, canc).ConfigureAwait(false);
#pragma warning disable HLQ012 // Loop body awaits; a CollectionsMarshal.AsSpan view cannot live across an await.
                    foreach (XlsSheetWriter sheet in _sheets)
                    {
                        frame.Reset();
                        BiffRecordWriter.WriteBof(frame, BiffRecord.SubstreamWorksheet);
                        BiffRecordWriter.WriteDimension(frame, sheet.RowCount, sheet.ColCount);
                        BiffRecordWriter.WriteWindow2(frame);
                        await dest.WriteAsync(frame.Memory, canc).ConfigureAwait(false);
                        await dest.WriteAsync(sheet.ColInfoMemory, canc).ConfigureAwait(false);
                        await dest.WriteAsync(sheet.CellsMemory, canc).ConfigureAwait(false);
                        frame.Reset();
                        BiffRecordWriter.WriteEof(frame);
                        await dest.WriteAsync(frame.Memory, canc).ConfigureAwait(false);
                    }
#pragma warning restore HLQ012
                }, ct).ConfigureAwait(false);
            }
            finally
            {
                globals.Dispose();
                frame.Dispose();
                foreach (ref readonly var sheet in CollectionsMarshal.AsSpan(_sheets))
                {
                    sheet.ReleaseBuffer();
                }
            }
        }

        /// <inheritdoc/>
        public ValueTask FlushAsync(CancellationToken ct = default)
        {
            ct.ThrowIfCancellationRequested();
            return new ValueTask(_stream.FlushAsync(ct));
        }

        /// <inheritdoc/>
        public async ValueTask DisposeAsync()
        {
            if (_disposed)
            {
                return;
            }
            _disposed = true;
            if (_state == WriterState.Started)
            {
                // Auto-finalize only a usable workbook; an unused/sheet-less one is abandoned
                // quietly so disposal never throws over cleanup.
                if (_sheets.Count > 0)
                {
                    await EndAsync().ConfigureAwait(false);
                }
                else
                {
                    _state = WriterState.Ended;
                }
            }
            if (!_leaveOpen)
            {
                await _stream.DisposeAsync().ConfigureAwait(false);
            }
        }
    }
}
