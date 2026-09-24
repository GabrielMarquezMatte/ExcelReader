using System.Globalization;
using System.IO.Compression;
using System.Security;
using System.Text;
using ExcelReader.Core.Writer.Internal;

namespace ExcelReader.Core.Writer.Xlsx
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
        private readonly List<(string Name, int SheetId, ExcelSheetVisibility Visibility)> _sheets = [];
        private bool _ended;
        private bool _sheetActive;
        private XlsxSheetWriter? _activeSheet;
        private readonly HashSet<string> _sheetNames = new(StringComparer.OrdinalIgnoreCase);
        private bool _faulted;
        private bool _zipClosed;
        private bool _disposed;

        private XlsxWorkbookWriter(ZipArchive zip, Stream stream, bool leaveOpen, XlsxWriterOptions options)
        {
            _zip = zip;
            _stream = stream;
            _leaveOpen = leaveOpen;
            UseSharedStrings = options.UseSharedStrings;
            _compression = options.Compression;
            _prefetchWrite = options.PrefetchWrite;
            _sharedStrings = options.UseSharedStrings ? new SharedStringTable() : null;
        }

        /// <summary>
        /// Creates a new <see cref="XlsxWorkbookWriter"/> that writes an XLSX package to <paramref name="stream"/>.
        /// </summary>
        /// <param name="stream">The destination stream; the returned writer takes ownership of the ZIP archive built on top of it.</param>
        /// <param name="leaveOpen">When <see langword="true"/>, <paramref name="stream"/> is left open after the workbook is disposed.</param>
        /// <param name="options">Compression, shared-string and background-deflate settings. Defaults to <see cref="XlsxWriterOptions.Default"/>.</param>
        /// <exception cref="ArgumentNullException"><paramref name="stream"/> is <see langword="null"/>.</exception>
        public static XlsxWorkbookWriter Create(Stream stream, bool leaveOpen = false, XlsxWriterOptions? options = null)
        {
            ArgumentNullException.ThrowIfNull(stream);
            ZipArchive zip = new(stream, ZipArchiveMode.Create, leaveOpen: true);
            return new XlsxWorkbookWriter(zip, stream, leaveOpen, options ?? XlsxWriterOptions.Default);
        }

        /// <inheritdoc/>
        /// <exception cref="ArgumentNullException"><paramref name="name"/> is <see langword="null"/>.</exception>
        /// <exception cref="ArgumentException"><paramref name="name"/> is empty, longer than 31 characters, or contains one of <c>: \ / ? * [ ]</c>.</exception>
        /// <exception cref="ObjectDisposedException">The workbook has already been ended.</exception>
        /// <exception cref="InvalidOperationException">The previously added sheet has not been ended, or a sheet named <paramref name="name"/> already exists.</exception>
        public XlsxSheetWriter AddSheet(string name)
        {
            return AddSheet(name, ExcelSheetVisibility.Visible);
        }

        /// <inheritdoc/>
        /// <exception cref="ArgumentNullException"><paramref name="name"/> is <see langword="null"/>.</exception>
        /// <exception cref="ArgumentException"><paramref name="name"/> is empty, longer than 31 characters, or contains one of <c>: \ / ? * [ ]</c>.</exception>
        /// <exception cref="ArgumentOutOfRangeException"><paramref name="visibility"/> is not a defined value.</exception>
        /// <exception cref="ObjectDisposedException">The workbook has already been ended.</exception>
        /// <exception cref="InvalidOperationException">The previously added sheet has not been ended, or a sheet named <paramref name="name"/> already exists.</exception>
        public XlsxSheetWriter AddSheet(string name, ExcelSheetVisibility visibility)
        {
            WriterStateGuard.RequireCanAddSheet(
                _ended, this, name, _sheetActive, nameof(XlsxSheetWriter), visibility);
            WriterStateGuard.ClaimSheetName(_sheetNames, name);
            _sheetActive = true;
            int sheetId = _sheets.Count + 1;
            _activeSheet = new XlsxSheetWriter(this, _zip, name, sheetId, visibility, _compression, _prefetchWrite);
            return _activeSheet;
        }

        internal void RegisterSheet(string name, int sheetId, ExcelSheetVisibility visibility)
        {
            _sheets.Add((name, sheetId, visibility));
        }

        internal void NotifySheetEnded(bool faulted)
        {
            _sheetActive = false;
            _activeSheet = null;
            if (faulted)
            {
                _ended = true;
                _faulted = true;
            }
        }

        internal bool UseSharedStrings { get; }

        internal int GetSharedStringIndex(string value)
        {
            return _sharedStrings!.GetOrAdd(value);
        }

        internal int GetSharedStringIndex(ReadOnlySpan<char> value)
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
        /// <exception cref="InvalidOperationException">No sheet has been added yet.</exception>
        public void End()
        {
            ObjectDisposedException.ThrowIf(_ended, this);
            _ended = true;
            try
            {
                _activeSheet?.Dispose();
                if (_sheets.Count == 0)
                {
                    throw new InvalidOperationException("An XLSX workbook must contain at least one sheet.");
                }
                WriteEntry("_rels/.rels", BuildRootRelsXml());
                WriteEntry("xl/styles.xml", BuildStylesXml());
                if (_sharedStrings is not null)
                {
                    using var bytes = _sharedStrings.ToXlsxBytes();
                    WriteEntry("xl/sharedStrings.xml", bytes.Memory.Span);
                }
                WriteEntry("xl/workbook.xml", BuildWorkbookXml());
                WriteEntry("xl/_rels/workbook.xml.rels", BuildWorkbookRelsXml());
                WriteEntry("[Content_Types].xml", BuildContentTypesXml());
                CloseZip();
            }
            catch
            {
                AbandonZip();
                throw;
            }
        }

        /// <inheritdoc/>
        /// <exception cref="ObjectDisposedException">The workbook has already been ended.</exception>
        /// <exception cref="InvalidOperationException">No sheet has been added yet.</exception>
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
                    throw new InvalidOperationException("An XLSX workbook must contain at least one sheet.");
                }
                await WriteEntryAsync("_rels/.rels", BuildRootRelsXml(), ct).ConfigureAwait(false);
                await WriteStylesAsync(ct).ConfigureAwait(false);
                if (_sharedStrings is not null)
                {
                    await WriteSharedStringsAsync(ct).ConfigureAwait(false);
                }
                await WriteWorkbookAsync(ct).ConfigureAwait(false);
                await WriteWorkbookRelsAsync(ct).ConfigureAwait(false);
                await WriteContentTypesAsync(ct).ConfigureAwait(false);
                await CloseZipAsync().ConfigureAwait(false);
            }
            catch
            {
                await AbandonZipAsync().ConfigureAwait(false);
                throw;
            }
        }

        private void CloseZip()
        {
            _zipClosed = true;
            _zip.Dispose();
        }

        private ValueTask CloseZipAsync()
        {
            _zipClosed = true;
            return _zip.DisposeAsync();
        }

        private void AbandonZip()
        {
            _faulted = true;
            if (!_zipClosed)
            {
                _zipClosed = true;
                FailureCleanup.Dispose(_zip);
            }
        }

        private ValueTask AbandonZipAsync()
        {
            _faulted = true;
            if (_zipClosed)
            {
                return ValueTask.CompletedTask;
            }
            _zipClosed = true;
            return FailureCleanup.DisposeAsync(_zip);
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
                    if (_sheets.Count == 0 && _activeSheet is null)
                    {
                        _ended = true;
                        CloseZip();
                    }
                    else
                    {
                        End();
                    }
                }
                else if (!_zipClosed)
                {
                    AbandonZip();
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
                    if (_sheets.Count == 0 && _activeSheet is null)
                    {
                        _ended = true;
                        await CloseZipAsync().ConfigureAwait(false);
                    }
                    else
                    {
                        await EndAsync().ConfigureAwait(false);
                    }
                }
                else if (!_zipClosed)
                {
                    await AbandonZipAsync().ConfigureAwait(false);
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

        private static string BuildRootRelsXml()
        {
            return "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                $"<Relationships xmlns=\"{XlsxConstants.PackageRelationshipsNs}\">" +
                $"<Relationship Id=\"rId1\" Type=\"{XlsxConstants.WorkbookRelType}\" Target=\"xl/workbook.xml\"/>" +
                "</Relationships>";
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

        private static string StateAttribute(ExcelSheetVisibility visibility)
        {
            return visibility switch
            {
                ExcelSheetVisibility.Hidden => " state=\"hidden\"",
                ExcelSheetVisibility.VeryHidden => " state=\"veryHidden\"",
                _ => "",
            };
        }

        private string BuildWorkbookXml()
        {
            StringBuilder sb = new();
            sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
            sb.Append($"<workbook xmlns=\"{XlsxConstants.MainNs}\" xmlns:r=\"{XlsxConstants.RelationshipsNs}\">");
            sb.Append("<sheets>");
            bool anyVisible = false;
            foreach ((string name, int sheetId, ExcelSheetVisibility visibility) in _sheets)
            {
                anyVisible |= visibility is ExcelSheetVisibility.Visible;
                sb.Append(CultureInfo.InvariantCulture,
                    $"<sheet name=\"{EscapeAttribute(name)}\" sheetId=\"{sheetId}\"{StateAttribute(visibility)} r:id=\"rId{sheetId + 1}\"/>");
            }
            WriterStateGuard.RequireVisibleSheet(anyVisible, nameof(XlsxWorkbookWriter));
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
            foreach ((_, int sheetId, _) in _sheets)
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
            foreach ((_, int sheetId, _) in _sheets)
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
