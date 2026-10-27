using System.Globalization;
using System.IO.Compression;
using System.Security;
using System.Text;
using ExcelReader.Core.Writer.Internal;
using ExcelReader.Core.Writer.Zip;

namespace ExcelReader.Core.Writer.Xlsx
{
    /// <summary>
    /// Writes a workbook to the XLSX (Office Open XML) format, streaming each sheet's rows into a ZIP archive as they're written.
    /// </summary>
    public sealed class XlsxWorkbookWriter : IWorkbookWriter<XlsxSheetWriter>, IZipWorkbookParts
    {
        private readonly ZipWorkbookPackage _package;
        private readonly CompressionLevel _compression;
        private readonly bool _prefetchWrite;
        private readonly SharedStringTable? _sharedStrings;
        private readonly StyleTable _styles = new();
        private readonly List<(string Name, int SheetId, ExcelSheetVisibility Visibility)> _sheets = [];
        private readonly List<(int SheetId, IReadOnlyList<WrittenTable> Tables)> _sheetTables = [];
        private readonly HashSet<string> _sheetNames = new(StringComparer.OrdinalIgnoreCase);

        private XlsxWorkbookWriter(Stream stream, bool leaveOpen, XlsxWriterOptions options)
        {
            _package = new ZipWorkbookPackage(this, "XLSX", stream, leaveOpen, options.Compression);
            UseSharedStrings = options.UseSharedStrings;
            Date1904 = options.Date1904;
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
            return new XlsxWorkbookWriter(stream, leaveOpen, options ?? XlsxWriterOptions.Default);
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
                _package.Ended, this, name, _package.ActiveSheet is not null, nameof(XlsxSheetWriter), visibility);
            WriterStateGuard.ClaimSheetName(_sheetNames, name);
            int sheetId = _sheets.Count + 1;
            var sheet = new XlsxSheetWriter(this, _package.Zip, name, sheetId, visibility, _compression, _prefetchWrite);
            _package.ActiveSheet = sheet;
            return sheet;
        }

        internal void RegisterSheet(string name, int sheetId, ExcelSheetVisibility visibility)
        {
            _sheets.Add((name, sheetId, visibility));
        }

        internal TableRegistry Tables { get; } = new();

        internal void RegisterTables(int sheetId, IReadOnlyList<WrittenTable> tables)
        {
            if (tables.Count > 0)
            {
                _sheetTables.Add((sheetId, tables));
            }
        }

        internal void NotifySheetEnded(bool faulted)
        {
            _package.SheetEnded(faulted);
        }

        internal bool UseSharedStrings { get; }

        internal bool Date1904 { get; }

        internal int GetSharedStringIndex(string value)
        {
            return _sharedStrings!.GetOrAdd(value);
        }

        internal int GetSharedStringIndex(ReadOnlySpan<char> value)
        {
            return _sharedStrings!.GetOrAdd(value);
        }

        internal SharedStringTable? SharedStrings
        {
            get
            {
                return _sharedStrings;
            }
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
            _package.End();
        }

        /// <inheritdoc/>
        /// <exception cref="ObjectDisposedException">The workbook has already been ended.</exception>
        /// <exception cref="InvalidOperationException">No sheet has been added yet.</exception>
        public ValueTask EndAsync(CancellationToken ct = default)
        {
            return _package.EndAsync(ct);
        }

        int IZipWorkbookParts.SheetCount
        {
            get
            {
                return _sheets.Count;
            }
        }

        void IZipWorkbookParts.WriteParts()
        {
            foreach ((int sheetId, IReadOnlyList<WrittenTable> tables) in _sheetTables)
            {
                _package.WriteEntry($"xl/worksheets/_rels/sheet{sheetId}.xml.rels", TablePackage.SheetRelsXml(tables, "xml"));
                foreach (WrittenTable table in tables)
                {
                    _package.WriteEntry($"xl/tables/table{table.Id}.xml", XlsxTableXml.Table(table));
                }
            }
            _package.WriteEntry("_rels/.rels", BuildRootRelsXml());
            _package.WriteEntry("xl/styles.xml", BuildStylesXml());
            if (_sharedStrings is not null)
            {
                _package.WriteSharedStrings("xl/sharedStrings.xml", _sharedStrings, binary: false);
            }
            _package.WriteEntry("xl/workbook.xml", BuildWorkbookXml());
            _package.WriteEntry("xl/_rels/workbook.xml.rels", BuildWorkbookRelsXml());
            _package.WriteEntry("[Content_Types].xml", BuildContentTypesXml());
        }

        async ValueTask IZipWorkbookParts.WritePartsAsync(CancellationToken ct)
        {
            foreach ((int sheetId, IReadOnlyList<WrittenTable> tables) in _sheetTables)
            {
                await _package.WriteEntryAsync($"xl/worksheets/_rels/sheet{sheetId}.xml.rels", TablePackage.SheetRelsXml(tables, "xml"), ct).ConfigureAwait(false);
                foreach (WrittenTable table in tables)
                {
                    await _package.WriteEntryAsync($"xl/tables/table{table.Id}.xml", XlsxTableXml.Table(table), ct).ConfigureAwait(false);
                }
            }
            await _package.WriteEntryAsync("_rels/.rels", BuildRootRelsXml(), ct).ConfigureAwait(false);
            await _package.WriteEntryAsync("xl/styles.xml", BuildStylesXml(), ct).ConfigureAwait(false);
            if (_sharedStrings is not null)
            {
                await _package.WriteSharedStringsAsync("xl/sharedStrings.xml", _sharedStrings, binary: false, ct).ConfigureAwait(false);
            }
            await _package.WriteEntryAsync("xl/workbook.xml", BuildWorkbookXml(), ct).ConfigureAwait(false);
            await _package.WriteEntryAsync("xl/_rels/workbook.xml.rels", BuildWorkbookRelsXml(), ct).ConfigureAwait(false);
            await _package.WriteEntryAsync("[Content_Types].xml", BuildContentTypesXml(), ct).ConfigureAwait(false);
        }

        /// <summary>
        /// Synchronous counterpart to <see cref="FlushAsync"/>, for native/unmanaged callers whose ABI
        /// is synchronous.
        /// </summary>
        public void Flush()
        {
            _package.Flush();
        }

        /// <inheritdoc/>
        public ValueTask FlushAsync(CancellationToken ct = default)
        {
            return _package.FlushAsync(ct);
        }

        /// <summary>
        /// Synchronous counterpart to <see cref="DisposeAsync"/>, for native/unmanaged callers whose ABI
        /// is synchronous.
        /// </summary>
        public void Dispose()
        {
            _package.Dispose();
        }

        /// <inheritdoc/>
        public ValueTask DisposeAsync()
        {
            return _package.DisposeAsync();
        }

        private static string BuildRootRelsXml()
        {
            return "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                $"<Relationships xmlns=\"{XlsxConstants.PackageRelationshipsNs}\">" +
                $"<Relationship Id=\"rId1\" Type=\"{XlsxConstants.WorkbookRelType}\" Target=\"xl/workbook.xml\"/>" +
                "</Relationships>";
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
            if (Date1904)
            {
                sb.Append("<workbookPr date1904=\"1\"/>");
            }
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
            foreach ((_, IReadOnlyList<WrittenTable> tables) in _sheetTables)
            {
                foreach (WrittenTable table in tables)
                {
                    sb.Append(CultureInfo.InvariantCulture, $"<Override PartName=\"/xl/tables/table{table.Id}.xml\" ContentType=\"{XlsxConstants.TableContentType}\"/>");
                }
            }
            sb.Append("</Types>");
            return sb.ToString();
        }

        private static string EscapeAttribute(string value)
        {
            return SecurityElement.Escape(value);
        }
    }
}
