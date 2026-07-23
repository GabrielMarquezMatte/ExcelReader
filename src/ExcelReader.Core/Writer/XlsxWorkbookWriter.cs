using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.IO.Compression;
using System.Security;
using System.Text;
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
        private readonly SharedStringTable? _sharedStrings;
        private readonly List<(string Name, int SheetId)> _sheets = [];
        private WriterState _state = WriterState.Created;
        private bool _sheetActive;
        private XlsxSheetWriter? _activeSheet;
        private bool _disposed;

        private XlsxWorkbookWriter(ZipArchive zip, Stream stream, bool leaveOpen, bool useSharedStrings, CompressionLevel compression)
        {
            _zip = zip;
            _stream = stream;
            _leaveOpen = leaveOpen;
            UseSharedStrings = useSharedStrings;
            _compression = compression;
            _sharedStrings = useSharedStrings ? new SharedStringTable() : null;
        }

        /// <summary>
        /// Creates a new <see cref="XlsxWorkbookWriter"/> that writes an XLSX package to <paramref name="stream"/>.
        /// </summary>
        /// <param name="stream">The destination stream; the returned writer takes ownership of the ZIP archive built on top of it.</param>
        /// <param name="leaveOpen">When <see langword="true"/>, <paramref name="stream"/> is left open after the workbook is disposed.</param>
        /// <param name="compression">The compression level applied to each ZIP entry.</param>
        /// <param name="useSharedStrings">When <see langword="true"/>, string cells are deduplicated into a shared string table instead of being inlined.</param>
        /// <param name="ct">A token to cancel the operation before the archive is created.</param>
        /// <exception cref="ArgumentNullException"><paramref name="stream"/> is <see langword="null"/>.</exception>
        [SuppressMessage("IDisposableAnalyzers.Correctness", "IDISP001:Dispose created",
            Justification = "Factory method transfers ZipArchive ownership to XlsxWorkbookWriter; caller disposes via DisposeAsync.")]
        [SuppressMessage("Reliability", "CA2000:Dispose objects before losing scope",
            Justification = "XlsxWorkbookWriter takes ownership of ZipArchive and disposes it in DisposeAsync/EndAsync.")]
        public static ValueTask<XlsxWorkbookWriter> CreateAsync(
            Stream stream, bool leaveOpen = false,
            CompressionLevel compression = CompressionLevel.Fastest,
            bool useSharedStrings = false,
            CancellationToken ct = default)
        {
            ArgumentNullException.ThrowIfNull(stream);
            ct.ThrowIfCancellationRequested();
            ZipArchive zip = new(stream, ZipArchiveMode.Create, leaveOpen: true);
            return ValueTask.FromResult(new XlsxWorkbookWriter(zip, stream, leaveOpen, useSharedStrings, compression));
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
        public XlsxSheetWriter AddSheet(string name)
        {
            ArgumentNullException.ThrowIfNull(name);
            WriterStateGuard.ThrowIfEnded(_state, this);
            WriterStateGuard.RequireStarted(_state, nameof(XlsxWorkbookWriter), "adding sheets");
            WriterStateGuard.ValidateSheetName(name);
            if (_sheetActive)
            {
                throw new InvalidOperationException("The previous XlsxSheetWriter must be ended before adding a new sheet.");
            }
            _sheetActive = true;
            int sheetId = _sheets.Count + 1;
            _activeSheet = new XlsxSheetWriter(this, _zip, name, sheetId, _compression);
            return _activeSheet;
        }

        internal void RegisterSheet(string name, int sheetId)
        {
            _sheets.Add((name, sheetId));
        }

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

        /// <inheritdoc/>
        public ValueTask FlushAsync(CancellationToken ct = default)
        {
            return ZipEntryWriter.FlushAsync(_stream, ct);
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
                // EndAsync deliberately rejects a zero-sheet workbook. Disposal still must release
                // a partially configured writer (for example after an earlier validation failure).
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

        private ValueTask WriteRootRelsAsync(CancellationToken ct)
        {
            return WriteEntryAsync(
                "_rels/.rels",
                "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                $"<Relationships xmlns=\"{XlsxConstants.PackageRelationshipsNs}\">" +
                $"<Relationship Id=\"rId1\" Type=\"{XlsxConstants.WorkbookRelType}\" Target=\"xl/workbook.xml\"/>" +
                "</Relationships>",
                ct);
        }

        private ValueTask WriteStylesAsync(CancellationToken ct)
        {
            return WriteEntryAsync(
                "xl/styles.xml",
                "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                $"<styleSheet xmlns=\"{XlsxConstants.MainNs}\">" +
                "<numFmts count=\"1\"><numFmt numFmtId=\"14\" formatCode=\"mm-dd-yy\"/></numFmts>" +
                "<fonts count=\"1\"><font/></fonts>" +
                "<fills count=\"2\"><fill><patternFill patternType=\"none\"/></fill><fill><patternFill patternType=\"gray125\"/></fill></fills>" +
                "<borders count=\"1\"><border/></borders>" +
                "<cellStyleXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/></cellStyleXfs>" +
                "<cellXfs count=\"2\">" +
                "<xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\"/>" +
                "<xf numFmtId=\"14\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyNumberFormat=\"1\"/>" +
                "</cellXfs>" +
                "</styleSheet>",
                ct);
        }

        private async ValueTask WriteSharedStringsAsync(CancellationToken ct)
        {
            using var bytes = _sharedStrings!.ToXlsxBytes();
            await WriteEntryAsync("xl/sharedStrings.xml", bytes.Memory, ct).ConfigureAwait(false);
        }

        private ValueTask WriteWorkbookAsync(CancellationToken ct)
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
            return WriteEntryAsync("xl/workbook.xml", sb.ToString(), ct);
        }

        private ValueTask WriteWorkbookRelsAsync(CancellationToken ct)
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
            return WriteEntryAsync("xl/_rels/workbook.xml.rels", sb.ToString(), ct);
        }

        private ValueTask WriteContentTypesAsync(CancellationToken ct)
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
            return WriteEntryAsync("[Content_Types].xml", sb.ToString(), ct);
        }

        private ValueTask WriteEntryAsync(string entryName, string content, CancellationToken ct)
        {
            return ZipEntryWriter.WriteTextAsync(_zip, entryName, content, _compression, ct);
        }

        private ValueTask WriteEntryAsync(string entryName, ReadOnlyMemory<byte> content, CancellationToken ct)
        {
            return ZipEntryWriter.WriteBytesAsync(_zip, entryName, content, _compression, ct);
        }

        private static string EscapeAttribute(string value)
        {
            return SecurityElement.Escape(value)!;
        }
    }
}
