using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.IO.Compression;
using System.Security;
using System.Text;
using ExcelReader.Core.Internal;
using ExcelReader.Core.Writer.Internal;

namespace ExcelReader.Core.Writer
{
    /// <summary>
    /// Writes a workbook to the XLSX (Office Open XML) format, streaming each sheet's rows into a ZIP archive as they're written.
    /// </summary>
    public sealed class XlsxWorkbookWriter : IWorkbookWriter<XlsxSheetWriter>
    {
        private readonly ZipArchive _zip;
        private readonly Stream _stream;
        private readonly bool _leaveOpen;
        private readonly CompressionLevel _compression;
        private readonly bool _prefetchWrite;
        private readonly SharedStringTable? _sharedStrings;
        private readonly StyleTable _styles = new();
        private readonly List<(string Name, int SheetId)> _sheets = [];
        private WriterState _state = WriterState.Created;
        private bool _sheetActive;
        private XlsxSheetWriter? _activeSheet;
        private bool _disposed;

        private XlsxWorkbookWriter(ZipArchive zip, Stream stream, bool leaveOpen, bool useSharedStrings, CompressionLevel compression, bool prefetchWrite)
        {
            _zip = zip;
            _stream = stream;
            _leaveOpen = leaveOpen;
            UseSharedStrings = useSharedStrings;
            _compression = compression;
            _prefetchWrite = prefetchWrite;
            _sharedStrings = useSharedStrings ? new SharedStringTable() : null;
        }

        /// <summary>
        /// Creates a new <see cref="XlsxWorkbookWriter"/> that writes an XLSX package to <paramref name="stream"/>.
        /// </summary>
        /// <param name="stream">The destination stream; the returned writer takes ownership of the ZIP archive built on top of it.</param>
        /// <param name="leaveOpen">When <see langword="true"/>, <paramref name="stream"/> is left open after the workbook is disposed.</param>
        /// <param name="compression">The compression level applied to each ZIP entry.</param>
        /// <param name="useSharedStrings">When <see langword="true"/>, string cells are deduplicated into a shared string table instead of being inlined.</param>
        /// <param name="prefetchWrite">
        /// When <see langword="true"/>, each sheet's deflate runs on a background thread instead of the
        /// calling thread, overlapping compression with row serialization. Defaults to <see langword="false"/>.
        /// Intended for single-file batch writing, where overlapping deflate with row-building shortens
        /// one write's wall-clock time — not for concurrent server workloads, where the extra background
        /// thread per sheet competes with work already saturating the CPU (mirrors
        /// <see cref="Reader.ExcelReaderOptions.PrefetchDecompression"/>'s own tradeoff on the read side).
        /// </param>
        /// <param name="ct">A token to cancel the operation before the archive is created.</param>
        /// <exception cref="ArgumentNullException"><paramref name="stream"/> is <see langword="null"/>.</exception>
        [SuppressMessage("IDisposableAnalyzers.Correctness", "IDISP001:Dispose created",
            Justification = "Factory method transfers ZipArchive ownership to XlsxWorkbookWriter; caller disposes via DisposeAsync.")]
        [SuppressMessage("IDisposableAnalyzers.Correctness", "IDISP004:Don't ignore created IDisposable",
            Justification = "Factory method transfers ZipArchive ownership to XlsxWorkbookWriter; caller disposes via DisposeAsync.")]
        [SuppressMessage("Reliability", "CA2000:Dispose objects before losing scope",
            Justification = "XlsxWorkbookWriter takes ownership of ZipArchive and disposes it in DisposeAsync/EndAsync.")]
        public static ValueTask<XlsxWorkbookWriter> CreateAsync(
            Stream stream, bool leaveOpen = false,
            CompressionLevel compression = CompressionLevel.Fastest,
            bool useSharedStrings = false,
            bool prefetchWrite = false,
            CancellationToken ct = default)
        {
            ArgumentNullException.ThrowIfNull(stream);
            ct.ThrowIfCancellationRequested();
            ZipArchive zip = new(stream, ZipArchiveMode.Create, leaveOpen: true);
            return ValueTask.FromResult(new XlsxWorkbookWriter(zip, stream, leaveOpen, useSharedStrings, compression, prefetchWrite));
        }

        /// <summary>
        /// Synchronous counterpart to <see cref="CreateAsync"/>, for native/unmanaged callers whose ABI
        /// is synchronous. Parameters mirror <see cref="CreateAsync"/> exactly, minus <c>ct</c>.
        /// </summary>
        [SuppressMessage("IDisposableAnalyzers.Correctness", "IDISP001:Dispose created",
            Justification = "Factory method transfers ZipArchive ownership to XlsxWorkbookWriter; caller disposes via Dispose/DisposeAsync.")]
        [SuppressMessage("IDisposableAnalyzers.Correctness", "IDISP004:Don't ignore created IDisposable",
            Justification = "Factory method transfers ZipArchive ownership to XlsxWorkbookWriter; caller disposes via Dispose/DisposeAsync.")]
        [SuppressMessage("Reliability", "CA2000:Dispose objects before losing scope",
            Justification = "XlsxWorkbookWriter takes ownership of ZipArchive and disposes it in Dispose(Async)/End(Async).")]
        public static XlsxWorkbookWriter Create(
            Stream stream, bool leaveOpen = false,
            CompressionLevel compression = CompressionLevel.Fastest,
            bool useSharedStrings = false,
            bool prefetchWrite = false)
        {
            ArgumentNullException.ThrowIfNull(stream);
            ZipArchive zip = new(stream, ZipArchiveMode.Create, leaveOpen: true);
            return new XlsxWorkbookWriter(zip, stream, leaveOpen, useSharedStrings, compression, prefetchWrite);
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
            WriterStateGuard.RequireCreated(_state, nameof(XlsxWorkbookWriter));
            _state = WriterState.Started;
            WriteRootRels();
        }

        /// <inheritdoc/>
        /// <exception cref="ObjectDisposedException">The workbook has already been ended.</exception>
        /// <exception cref="InvalidOperationException">The workbook has already been started.</exception>
        public ValueTask StartAsync(CancellationToken ct = default)
        {
            WriterStateGuard.ThrowIfEnded(_state, this);
            WriterStateGuard.RequireCreated(_state, nameof(XlsxWorkbookWriter));
            ct.ThrowIfCancellationRequested();
            _state = WriterState.Started;
            return WriteRootRelsAsync(ct);
        }

        /// <inheritdoc/>
        /// <exception cref="ArgumentNullException"><paramref name="name"/> is <see langword="null"/>.</exception>
        /// <exception cref="ArgumentException"><paramref name="name"/> is empty, longer than 31 characters, or contains one of <c>: \ / ? * [ ]</c>.</exception>
        /// <exception cref="ObjectDisposedException">The workbook has already been ended.</exception>
        /// <exception cref="InvalidOperationException">The workbook has not been started, or the previously added sheet has not been ended.</exception>
        [SuppressMessage("IDisposableAnalyzers.Correctness", "IDISP003:Dispose previous before re-assigning",
            Justification = "RequireCanAddSheet above guarantees _activeSheet is null (any previous sheet was already ended and disposed) before this assigns a new one.")]
        public XlsxSheetWriter AddSheet(string name)
        {
            WriterStateGuard.RequireCanAddSheet(
                _state, this, nameof(XlsxWorkbookWriter), name, _sheetActive, nameof(XlsxSheetWriter));
            _sheetActive = true;
            int sheetId = _sheets.Count + 1;
            _activeSheet = new XlsxSheetWriter(this, _zip, name, sheetId, _compression, _prefetchWrite);
            return _activeSheet;
        }

        internal void RegisterSheet(string name, int sheetId)
        {
            _sheets.Add((name, sheetId));
        }

        [SuppressMessage("IDisposableAnalyzers.Correctness", "IDISP003:Dispose previous before re-assigning",
            Justification = "Called only after the active sheet's own End/Dispose has already run (from End/Dispose below); this just clears the now-stale reference.")]
        internal void NotifySheetEnded()
        {
            _sheetActive = false;
            _activeSheet = null;
        }

        internal bool UseSharedStrings { get; }

        internal int GetSharedStringIndex(string value)
        {
            return _sharedStrings!.GetOrAdd(value);
        }

        /// <inheritdoc/>
        public int AddStyle(CellStyle style)
        {
            return _styles.Add(style);
        }

        internal int StyleCount => _styles.Count;

        /// <summary>
        /// Synchronous counterpart to <see cref="EndAsync"/>, for native/unmanaged callers whose ABI is
        /// synchronous.
        /// </summary>
        /// <exception cref="ObjectDisposedException">The workbook has already been ended.</exception>
        /// <exception cref="InvalidOperationException">The workbook has not been started, or no sheet has been added yet.</exception>
        public void End()
        {
            WriterStateGuard.ThrowIfEnded(_state, this);
            WriterStateGuard.RequireStarted(_state, nameof(XlsxWorkbookWriter), "ending");
            if (_sheets.Count == 0)
            {
                throw new InvalidOperationException("An XLSX workbook must contain at least one sheet.");
            }
            _state = WriterState.Ended;
            if (_activeSheet is not null)
            {
                _activeSheet.Dispose();
            }
            WriteEntry("xl/styles.xml", BuildStylesXml());
            if (_sharedStrings is not null)
            {
                using var bytes = _sharedStrings.ToXlsxBytes();
                WriteEntry("xl/sharedStrings.xml", bytes.Memory.Span);
            }
            WriteEntry("xl/workbook.xml", BuildWorkbookXml());
            WriteEntry("xl/_rels/workbook.xml.rels", BuildWorkbookRelsXml());
            WriteEntry("[Content_Types].xml", BuildContentTypesXml());
            ZipArchiveDisposal.Dispose(_zip);
        }

        /// <inheritdoc/>
        /// <exception cref="ObjectDisposedException">The workbook has already been ended.</exception>
        /// <exception cref="InvalidOperationException">The workbook has not been started, or no sheet has been added yet.</exception>
        public async ValueTask EndAsync(CancellationToken ct = default)
        {
            WriterStateGuard.ThrowIfEnded(_state, this);
            WriterStateGuard.RequireStarted(_state, nameof(XlsxWorkbookWriter), "ending");
            ct.ThrowIfCancellationRequested();
            if (_sheets.Count == 0)
            {
                throw new InvalidOperationException("An XLSX workbook must contain at least one sheet.");
            }
            _state = WriterState.Ended;
            if (_activeSheet is not null)
            {
                await _activeSheet.DisposeAsync().ConfigureAwait(false);
            }
            await WriteStylesAsync(ct).ConfigureAwait(false);
            if (_sharedStrings is not null)
            {
                await WriteSharedStringsAsync(ct).ConfigureAwait(false);
            }
            await WriteWorkbookAsync(ct).ConfigureAwait(false);
            await WriteWorkbookRelsAsync(ct).ConfigureAwait(false);
            await WriteContentTypesAsync(ct).ConfigureAwait(false);
            await ZipArchiveDisposal.DisposeAsync(_zip).ConfigureAwait(false);
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
            return ZipEntryWriter.FlushAsync(_stream, ct);
        }

        /// <summary>
        /// Synchronous counterpart to <see cref="DisposeAsync"/>, for native/unmanaged callers whose ABI
        /// is synchronous.
        /// </summary>
        [SuppressMessage("IDisposableAnalyzers.Correctness", "IDISP021:Call this.Dispose(true)",
            Justification = "Sealed type, no finalizer, no Dispose(bool) pattern to route through — this is the only Dispose overload.")]
        public void Dispose()
        {
            if (_disposed)
            {
                return;
            }
            _disposed = true;
            if (_state == WriterState.Started)
            {
                // End rejects a zero-sheet workbook; disposal must still release a partial writer.
                if (_sheets.Count == 0)
                {
                    _state = WriterState.Ended;
                    ZipArchiveDisposal.Dispose(_zip);
                }
                else
                {
                    End();
                }
            }
            else if (_state == WriterState.Created)
            {
                ZipArchiveDisposal.Dispose(_zip);
            }
            if (!_leaveOpen)
            {
                _stream.Dispose();
            }
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
                // EndAsync rejects a zero-sheet workbook; disposal must still release a partial writer.
                if (_sheets.Count == 0)
                {
                    _state = WriterState.Ended;
                    await ZipArchiveDisposal.DisposeAsync(_zip).ConfigureAwait(false);
                }
                else
                {
                    await EndAsync().ConfigureAwait(false);
                }
            }
            else if (_state == WriterState.Created)
            {
                await ZipArchiveDisposal.DisposeAsync(_zip).ConfigureAwait(false);
            }
            if (!_leaveOpen)
            {
                await _stream.DisposeAsync().ConfigureAwait(false);
            }
        }

        private static string BuildRootRelsXml()
        {
            return "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                $"<Relationships xmlns=\"{XlsxConstants.PackageRelationshipsNs}\">" +
                $"<Relationship Id=\"rId1\" Type=\"{XlsxConstants.WorkbookRelType}\" Target=\"xl/workbook.xml\"/>" +
                "</Relationships>";
        }

        private void WriteRootRels()
        {
            WriteEntry("_rels/.rels", BuildRootRelsXml());
        }

        private ValueTask WriteRootRelsAsync(CancellationToken ct)
        {
            return WriteEntryAsync("_rels/.rels", BuildRootRelsXml(), ct);
        }

        private ValueTask WriteStylesAsync(CancellationToken ct)
        {
            return WriteEntryAsync("xl/styles.xml", BuildStylesXml(), ct);
        }

        private string BuildStylesXml()
        {
            IReadOnlyList<CellStyle> styles = _styles.Styles;
            Dictionary<string, int> numFmtIds = _styles.AssignCustomNumberFormatIds();
            Dictionary<(bool Bold, bool Italic), int> fontIds = _styles.AssignFontIds();
            var fontsByIndex = new (bool Bold, bool Italic)[fontIds.Count];
            foreach (var (key, index) in fontIds)
            {
                fontsByIndex[index] = key;
            }

            var sb = new StringBuilder();
            sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
            sb.Append(CultureInfo.InvariantCulture, $"<styleSheet xmlns=\"{XlsxConstants.MainNs}\">");
            AppendNumFmts(sb, numFmtIds);
            AppendFonts(sb, fontsByIndex);
            sb.Append("<fills count=\"2\"><fill><patternFill patternType=\"none\"/></fill><fill><patternFill patternType=\"gray125\"/></fill></fills>");
            sb.Append("<borders count=\"1\"><border/></borders>");
            sb.Append("<cellStyleXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/></cellStyleXfs>");
            AppendCellXfs(sb, styles, numFmtIds, fontIds);
            sb.Append("</styleSheet>");
            return sb.ToString();
        }

        private static void AppendNumFmts(StringBuilder sb, Dictionary<string, int> numFmtIds)
        {
            sb.Append(CultureInfo.InvariantCulture, $"<numFmts count=\"{1 + numFmtIds.Count}\">");
            sb.Append("<numFmt numFmtId=\"14\" formatCode=\"mm-dd-yy\"/>");
            // Ordered by id: Dictionary enumeration order isn't guaranteed, but this XML must be
            // deterministic across runs.
            foreach (KeyValuePair<string, int> entry in numFmtIds.OrderBy(static kv => kv.Value))
            {
                sb.Append(CultureInfo.InvariantCulture, $"<numFmt numFmtId=\"{entry.Value}\" formatCode=\"{EscapeAttribute(entry.Key)}\"/>");
            }
            sb.Append("</numFmts>");
        }

        private static void AppendFonts(StringBuilder sb, (bool Bold, bool Italic)[] fontsByIndex)
        {
            sb.Append(CultureInfo.InvariantCulture, $"<fonts count=\"{fontsByIndex.Length}\">");
            foreach ((bool bold, bool italic) in fontsByIndex)
            {
                if (!bold && !italic)
                {
                    sb.Append("<font/>");
                    continue;
                }
                sb.Append("<font>");
                if (bold)
                {
                    sb.Append("<b/>");
                }
                if (italic)
                {
                    sb.Append("<i/>");
                }
                sb.Append("</font>");
            }
            sb.Append("</fonts>");
        }

        private static void AppendCellXfs(StringBuilder sb, IReadOnlyList<CellStyle> styles,
            Dictionary<string, int> numFmtIds, Dictionary<(bool Bold, bool Italic), int> fontIds)
        {
            sb.Append(CultureInfo.InvariantCulture, $"<cellXfs count=\"{styles.Count}\">");
            sb.Append("<xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\"/>");
            sb.Append("<xf numFmtId=\"14\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyNumberFormat=\"1\"/>");
            for (int i = 2; i < styles.Count; i++)
            {
                CellStyle style = styles[i];
                int numFmtId = style.NumberFormat is not null ? numFmtIds[style.NumberFormat] : 0;
                int fontId = fontIds[(style.Bold, style.Italic)];
                sb.Append(CultureInfo.InvariantCulture, $"<xf numFmtId=\"{numFmtId}\" fontId=\"{fontId}\" fillId=\"0\" borderId=\"0\" xfId=\"0\"");
                if (style.NumberFormat is not null)
                {
                    sb.Append(" applyNumberFormat=\"1\"");
                }
                if (style.Bold || style.Italic)
                {
                    sb.Append(" applyFont=\"1\"");
                }
                sb.Append("/>");
            }
            sb.Append("</cellXfs>");
        }

        private async ValueTask WriteSharedStringsAsync(CancellationToken ct)
        {
            using var bytes = _sharedStrings!.ToXlsxBytes();
            await WriteEntryAsync("xl/sharedStrings.xml", bytes.Memory, ct).ConfigureAwait(false);
        }

        private string BuildWorkbookXml()
        {
            StringBuilder sb = new();
            sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
            sb.Append($"<workbook xmlns=\"{XlsxConstants.MainNs}\" xmlns:r=\"{XlsxConstants.RelationshipsNs}\">");
            sb.Append("<sheets>");
            foreach ((string name, int sheetId) in _sheets)
            {
                sb.Append(CultureInfo.InvariantCulture, $"<sheet name=\"{EscapeAttribute(name)}\" sheetId=\"{sheetId}\" r:id=\"rId{sheetId + 1}\"/>");
            }
            sb.Append("</sheets></workbook>");
            return sb.ToString();
        }

        private ValueTask WriteWorkbookAsync(CancellationToken ct)
        {
            return WriteEntryAsync("xl/workbook.xml", BuildWorkbookXml(), ct);
        }

        private string BuildWorkbookRelsXml()
        {
            StringBuilder sb = new();
            sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
            sb.Append($"<Relationships xmlns=\"{XlsxConstants.PackageRelationshipsNs}\">");
            sb.Append($"<Relationship Id=\"rId1\" Type=\"{XlsxConstants.StylesRelType}\" Target=\"styles.xml\"/>");
            if (UseSharedStrings)
            {
                sb.Append($"<Relationship Id=\"rIdShared\" Type=\"{XlsxConstants.SharedStringsRelType}\" Target=\"sharedStrings.xml\"/>");
            }
            foreach ((_, int sheetId) in _sheets)
            {
                sb.Append(CultureInfo.InvariantCulture, $"<Relationship Id=\"rId{sheetId + 1}\" Type=\"{XlsxConstants.WorksheetRelType}\" Target=\"worksheets/sheet{sheetId}.xml\"/>");
            }
            sb.Append("</Relationships>");
            return sb.ToString();
        }

        private ValueTask WriteWorkbookRelsAsync(CancellationToken ct)
        {
            return WriteEntryAsync("xl/_rels/workbook.xml.rels", BuildWorkbookRelsXml(), ct);
        }

        private string BuildContentTypesXml()
        {
            StringBuilder sb = new();
            sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
            sb.Append($"<Types xmlns=\"{XlsxConstants.ContentTypesNs}\">");
            sb.Append($"<Default Extension=\"rels\" ContentType=\"{XlsxConstants.RelationshipsContentType}\"/>");
            sb.Append("<Default Extension=\"xml\" ContentType=\"application/xml\"/>");
            sb.Append($"<Override PartName=\"/xl/workbook.xml\" ContentType=\"{XlsxConstants.WorkbookContentType}\"/>");
            sb.Append($"<Override PartName=\"/xl/styles.xml\" ContentType=\"{XlsxConstants.StylesContentType}\"/>");
            if (UseSharedStrings)
            {
                sb.Append($"<Override PartName=\"/xl/sharedStrings.xml\" ContentType=\"{XlsxConstants.SharedStringsContentType}\"/>");
            }
            foreach ((_, int sheetId) in _sheets)
            {
                sb.Append(CultureInfo.InvariantCulture, $"<Override PartName=\"/xl/worksheets/sheet{sheetId}.xml\" ContentType=\"{XlsxConstants.WorksheetContentType}\"/>");
            }
            sb.Append("</Types>");
            return sb.ToString();
        }

        private ValueTask WriteContentTypesAsync(CancellationToken ct)
        {
            return WriteEntryAsync("[Content_Types].xml", BuildContentTypesXml(), ct);
        }

        private ValueTask WriteEntryAsync(string entryName, string content, CancellationToken ct)
        {
            return ZipEntryWriter.WriteTextAsync(_zip, entryName, content, _compression, ct);
        }

        private ValueTask WriteEntryAsync(string entryName, ReadOnlyMemory<byte> content, CancellationToken ct)
        {
            return ZipEntryWriter.WriteBytesAsync(_zip, entryName, content, _compression, ct);
        }

        private void WriteEntry(string entryName, string content)
        {
            ZipEntryWriter.WriteText(_zip, entryName, content, _compression);
        }

        private void WriteEntry(string entryName, ReadOnlySpan<byte> content)
        {
            ZipEntryWriter.WriteBytes(_zip, entryName, content, _compression);
        }

        private static string EscapeAttribute(string value)
        {
            return SecurityElement.Escape(value);
        }
    }
}
