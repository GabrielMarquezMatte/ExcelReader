using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;
using ExcelReader.Core.Writer.Internal;
using ExcelReader.Core.Writer.Xlsx;

namespace ExcelReader.Core.Writer.Xls
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
        [SuppressMessage("Usage", "CA2213:Disposable fields should be disposed",
            Justification = "Released with the sheet buffers once End or EndAsync has written the workbook.")]
        private readonly XlsSharedStrings _sharedStrings = new();
        private readonly List<XlsSheetWriter> _sheets = [];
        private bool _ended;
        private XlsSheetWriter? _activeSheet;
        private readonly HashSet<string> _sheetNames = new(StringComparer.OrdinalIgnoreCase);
        private bool _faulted;
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
        /// <param name="options">The date system; <see cref="XlsWriterOptions.Default"/> when <see langword="null"/>.</param>
        public static XlsWorkbookWriter Create(Stream stream, bool leaveOpen = false, XlsWriterOptions? options = null)
        {
            ArgumentNullException.ThrowIfNull(stream);
            return new XlsWorkbookWriter(stream, leaveOpen, (options ?? XlsWriterOptions.Default).Date1904);
        }

        /// <inheritdoc/>
        /// <exception cref="InvalidOperationException">The previously added sheet has not been ended yet.</exception>
        public XlsSheetWriter AddSheet(string name)
        {
            return AddSheet(name, ExcelSheetVisibility.Visible);
        }

        /// <inheritdoc/>
        /// <exception cref="InvalidOperationException">The previously added sheet has not been ended yet.</exception>
        public XlsSheetWriter AddSheet(string name, ExcelSheetVisibility visibility)
        {
            WriterStateGuard.RequireCanAddSheet(
                _ended, this, name, _activeSheet is not null, nameof(XlsSheetWriter), visibility);
            WriterStateGuard.ClaimSheetName(_sheetNames, name);
#pragma warning disable IDISP003
            _activeSheet = new XlsSheetWriter(this, name, _date1904, visibility);
#pragma warning restore IDISP003
            return _activeSheet;
        }

        internal int SheetCount => _sheets.Count;

        internal bool TryClaimSheetName(string name)
        {
            return _sheetNames.Add(name);
        }

        /// <inheritdoc/>
        public int AddStyle(CellStyle style)
        {
            return _styles.Add(style);
        }

        internal int StyleCount => _styles.Count;

        internal int DateStyle(int styleId) => _styles.DateVariant(styleId);

        internal int AddSharedString(ReadOnlySpan<char> value)
        {
            return _sharedStrings.Add(value);
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

        private BiffBuffer BuildGlobals(List<XlsSheetLayout> layouts, out int workbookSize)
        {
            string[] names = new string[_sheets.Count];
            ExcelSheetVisibility[] visibilities = new ExcelSheetVisibility[_sheets.Count];
            int activeSheet = -1;
            for (int i = 0; i < _sheets.Count; i++)
            {
                names[i] = _sheets[i].Name;
                visibilities[i] = _sheets[i].Visibility;
                if (activeSheet < 0 && visibilities[i] is ExcelSheetVisibility.Visible)
                {
                    activeSheet = i;
                }
            }
            WriterStateGuard.RequireVisibleSheet(activeSheet >= 0, nameof(XlsWorkbookWriter));

            for (int i = 0; i < _sheets.Count; i++)
            {
                layouts.Add(new XlsSheetLayout(_sheets[i], selected: i == activeSheet));
            }

            BiffBuffer globals = new(4096 + _sharedStrings.EstimatedLength);
            try
            {
                int[] offsetPositions = XlsGlobals.Write(globals, names, visibilities, activeSheet, _date1904, _styles, _sharedStrings);
                int offset = globals.Length;
                for (int i = 0; i < layouts.Count; i++)
                {
                    globals.PatchI32(offsetPositions[i], offset);
                    layouts[i].Place(offset);
                    offset += layouts[i].Length;
                }
                workbookSize = offset;
                return globals;
            }
            catch
            {
                globals.Dispose();
                throw;
            }
        }

        private void ReleaseStream()
        {
            if (_leaveOpen)
            {
                return;
            }
            if (_faulted)
            {
                FailureCleanup.Dispose(_stream);
                return;
            }
            _stream.Dispose();
        }

        private ValueTask ReleaseStreamAsync()
        {
            if (_leaveOpen)
            {
                return ValueTask.CompletedTask;
            }
            return _faulted ? FailureCleanup.DisposeAsync(_stream) : _stream.DisposeAsync();
        }

        private static void DisposeLayouts(List<XlsSheetLayout> layouts)
        {
            foreach (XlsSheetLayout layout in layouts)
            {
                layout.Dispose();
            }
        }

        private void ReleaseSheetBuffers()
        {
            _sharedStrings.Dispose();
            foreach (ref readonly var sheet in CollectionsMarshal.AsSpan(_sheets))
            {
                sheet.ReleaseBuffer();
            }
        }

        /// <summary>
        /// Synchronous counterpart to <see cref="EndAsync"/>, for native/unmanaged callers whose ABI is
        /// synchronous.
        /// </summary>
        /// <exception cref="InvalidOperationException">No sheet was ever added to the workbook.</exception>
        public void End()
        {
            ObjectDisposedException.ThrowIf(_ended, this);
            _ended = true;
            try
            {
                _activeSheet?.Dispose();
                if (_sheets.Count == 0)
                {
                    throw new InvalidOperationException("A workbook must contain at least one sheet.");
                }

                List<XlsSheetLayout> layouts = [];
                try
                {
                    using BiffBuffer globals = BuildGlobals(layouts, out int workbookSize);
                    OleCompoundWriter.Write(_stream, workbookSize, dest =>
                    {
                        dest.Write(globals.Span);
                        foreach (XlsSheetLayout layout in layouts)
                        {
                            layout.WriteTo(dest);
                        }
                    });
                }
                finally
                {
                    DisposeLayouts(layouts);
                }
            }
            catch
            {
                _faulted = true;
                throw;
            }
            finally
            {
                ReleaseSheetBuffers();
            }
        }

        /// <inheritdoc/>
        /// <exception cref="InvalidOperationException">No sheet was ever added to the workbook.</exception>
        public async ValueTask EndAsync(CancellationToken ct = default)
        {
            ObjectDisposedException.ThrowIf(_ended, this);
            ct.ThrowIfCancellationRequested();
            _ended = true;
            try
            {
                if (_activeSheet is not null)
                {
                    await _activeSheet.DisposeAsync().ConfigureAwait(false);
                }
                if (_sheets.Count == 0)
                {
                    throw new InvalidOperationException("A workbook must contain at least one sheet.");
                }

                List<XlsSheetLayout> layouts = [];
                try
                {
                    using BiffBuffer globals = BuildGlobals(layouts, out int workbookSize);
                    await OleCompoundWriter.WriteAsync(_stream, workbookSize, async (dest, canc) =>
                    {
                        await dest.WriteAsync(globals.Memory, canc).ConfigureAwait(false);
                        foreach (XlsSheetLayout layout in layouts)
                        {
                            await layout.WriteToAsync(dest, canc).ConfigureAwait(false);
                        }
                    }, ct).ConfigureAwait(false);
                }
                finally
                {
                    DisposeLayouts(layouts);
                }
            }
            catch
            {
                _faulted = true;
                throw;
            }
            finally
            {
                ReleaseSheetBuffers();
            }
        }

        /// <summary>
        /// Synchronous counterpart to <see cref="FlushAsync"/>, for native/unmanaged callers whose ABI
        /// is synchronous.
        /// </summary>
        public void Flush()
        {
            _stream.Flush();
        }

        /// <inheritdoc/>
        public ValueTask FlushAsync(CancellationToken ct = default)
        {
            ct.ThrowIfCancellationRequested();
            return new ValueTask(_stream.FlushAsync(ct));
        }

        /// <summary>
        /// Synchronous counterpart to <see cref="DisposeAsync"/>, for native/unmanaged callers whose ABI
        /// is synchronous.
        /// </summary>
        public void Dispose()
        {
            if (_disposed)
            {
                return;
            }
            _disposed = true;
            try
            {
                if (!_ended)
                {
                    if (_sheets.Count > 0 || _activeSheet is not null)
                    {
                        End();
                    }
                    else
                    {
                        _ended = true;
                    }
                }
            }
            catch
            {
                _faulted = true;
                ReleaseStream();
                throw;
            }
            ReleaseStream();
        }

        /// <inheritdoc/>
        public async ValueTask DisposeAsync()
        {
            if (_disposed)
            {
                return;
            }
            _disposed = true;
            try
            {
                if (!_ended)
                {
                    if (_sheets.Count > 0 || _activeSheet is not null)
                    {
                        await EndAsync().ConfigureAwait(false);
                    }
                    else
                    {
                        _ended = true;
                    }
                }
            }
            catch
            {
                _faulted = true;
                await ReleaseStreamAsync().ConfigureAwait(false);
                throw;
            }
            await ReleaseStreamAsync().ConfigureAwait(false);
        }
    }
}
