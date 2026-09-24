using System.Globalization;
using System.IO.Compression;
using System.Runtime.InteropServices;
using System.Text;
using ExcelReader.Core.Reader.Xlsb;
using ExcelReader.Core.Writer.Internal;

namespace ExcelReader.Core.Writer.Xlsb
{
    /// <summary>Writes an .xlsb workbook (BIFF12 binary format) to a stream, one sheet at a time.</summary>
    public sealed class XlsbWorkbookWriter : IWorkbookWriter<XlsbSheetWriter>
    {
        private readonly ZipArchive _zip;
        private readonly Stream _stream;
        private readonly bool _leaveOpen;
        private readonly bool _date1904;
        private readonly CompressionLevel _compression;
        private readonly bool _prefetchWrite;
        private readonly SharedStringTable? _sharedStrings;
        private readonly StyleTable _styles = new();
        private readonly List<XlsbSheetWriter> _sheets = [];
        private bool _ended;
        private XlsbSheetWriter? _activeSheet;
        private readonly HashSet<string> _sheetNames = new(StringComparer.OrdinalIgnoreCase);
        private bool _faulted;
        private bool _zipClosed;
        private bool _disposed;

        private XlsbWorkbookWriter(ZipArchive zip, Stream stream, bool leaveOpen, XlsbWriterOptions options)
        {
            _zip = zip;
            _stream = stream;
            _leaveOpen = leaveOpen;
            _date1904 = options.Date1904;
            UseSharedStrings = options.UseSharedStrings;
            _compression = options.Compression;
            _prefetchWrite = options.PrefetchWrite;
            _sharedStrings = options.UseSharedStrings ? new SharedStringTable() : null;
        }

        /// <summary>Creates a writer that produces an .xlsb archive on <paramref name="stream"/>.</summary>
        /// <param name="stream">The destination stream; must be writable.</param>
        /// <param name="leaveOpen">If <see langword="true"/>, <paramref name="stream"/> is not disposed when the writer is disposed.</param>
        /// <param name="options">Date system, compression, shared-string and background-deflate settings. Defaults to <see cref="XlsbWriterOptions.Default"/>.</param>
        /// <exception cref="ArgumentNullException"><paramref name="stream"/> is <see langword="null"/>.</exception>
        public static XlsbWorkbookWriter Create(Stream stream, bool leaveOpen = false, XlsbWriterOptions? options = null)
        {
            ArgumentNullException.ThrowIfNull(stream);
            ZipArchive zip = new(stream, ZipArchiveMode.Create, leaveOpen: true);
            return new XlsbWorkbookWriter(zip, stream, leaveOpen, options ?? XlsbWriterOptions.Default);
        }

        /// <inheritdoc/>
        public XlsbSheetWriter AddSheet(string name)
        {
            return AddSheet(name, ExcelSheetVisibility.Visible);
        }

        /// <inheritdoc/>
        public XlsbSheetWriter AddSheet(string name, ExcelSheetVisibility visibility)
        {
            WriterStateGuard.RequireCanAddSheet(
                _ended, this, name, _activeSheet is not null, nameof(XlsbSheetWriter), visibility);
            WriterStateGuard.ClaimSheetName(_sheetNames, name);
            int sheetId = _sheets.Count + 1;
            _activeSheet = new XlsbSheetWriter(this, _zip, name, sheetId, visibility, _date1904, _compression, _prefetchWrite);
            return _activeSheet;
        }

        internal void RegisterSheet(XlsbSheetWriter sheet)
        {
            _sheets.Add(sheet);
        }

        internal void NotifySheetEnded(bool faulted)
        {
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

                WriteRootRels();
                WriteWorkbook();
                WriteWorkbookRels();
                WriteStyles();
                WriteSharedStrings();
                WriteAppProperties();
                WriteContentTypes();
                CloseZip();
            }
            catch
            {
                AbandonZip();
                throw;
            }
        }

        /// <inheritdoc/>
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

                await WriteRootRelsAsync(ct).ConfigureAwait(false);
                await WriteWorkbookAsync(ct).ConfigureAwait(false);
                await WriteWorkbookRelsAsync(ct).ConfigureAwait(false);
                await WriteStylesAsync(ct).ConfigureAwait(false);
                await WriteSharedStringsAsync(ct).ConfigureAwait(false);
                await WriteAppPropertiesAsync(ct).ConfigureAwait(false);
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
                "<Relationship Id=\"app\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties\" Target=\"docProps/app.xml\"/>" +
                $"<Relationship Id=\"wb\" Type=\"{XlsxConstants.WorkbookRelType}\" Target=\"xl/workbook.bin\"/>" +
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

        private void BuildWorkbookBin(BiffBuffer payload, BiffBuffer data)
        {
            Biff12RecordWriter.WriteRecord(data, Brt.BeginBook);
            payload.WriteU32(_date1904 ? 1u : 0u);
            payload.WriteU32(0);
            Biff12RecordWriter.WriteWideString(payload, string.Empty);
            Biff12RecordWriter.WriteRecord(data, Brt.WbProp, payload.Span);

            Biff12RecordWriter.WriteRecord(data, Brt.BeginBundleShs);
            bool anyVisible = false;
            foreach (ref readonly var sheet in CollectionsMarshal.AsSpan(_sheets))
            {
                anyVisible |= sheet.Visibility is ExcelSheetVisibility.Visible;
                payload.Reset();
                payload.WriteU32((uint)sheet.Visibility);
                payload.WriteU32((uint)sheet.SheetId);
                Biff12RecordWriter.WriteWideString(payload, $"s{sheet.SheetId}");
                Biff12RecordWriter.WriteWideString(payload, sheet.Name);
                Biff12RecordWriter.WriteRecord(data, Brt.BundleSh, payload.Span);
            }
            WriterStateGuard.RequireVisibleSheet(anyVisible, nameof(XlsbWorkbookWriter));
            Biff12RecordWriter.WriteRecord(data, Brt.EndBundleShs);
            Biff12RecordWriter.WriteRecord(data, Brt.EndBook);
        }

        private void WriteWorkbook()
        {
            using BiffBuffer payload = new(256);
            using BiffBuffer data = new(1024);
            BuildWorkbookBin(payload, data);
            WriteEntry("xl/workbook.bin", data.Span);
        }

        private async ValueTask WriteWorkbookAsync(CancellationToken ct)
        {
            using BiffBuffer payload = new(256);
            using BiffBuffer data = new(1024);
            BuildWorkbookBin(payload, data);
            await WriteEntryAsync("xl/workbook.bin", data.Memory, ct).ConfigureAwait(false);
        }

        private string BuildWorkbookRelsXml()
        {
            StringBuilder sb = new();
            sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
            sb.Append(CultureInfo.InvariantCulture, $"<Relationships xmlns=\"{XlsxConstants.PackageRelationshipsNs}\">");
            sb.Append(CultureInfo.InvariantCulture, $"<Relationship Id=\"s\" Type=\"{XlsxConstants.StylesRelType}\" Target=\"styles.bin\"/>");
            sb.Append(CultureInfo.InvariantCulture, $"<Relationship Id=\"ss\" Type=\"{XlsxConstants.SharedStringsRelType}\" Target=\"sharedStrings.bin\"/>");
            foreach (ref readonly var sheet in CollectionsMarshal.AsSpan(_sheets))
            {
                sb.Append(CultureInfo.InvariantCulture, $"<Relationship Id=\"s{sheet.SheetId}\" Type=\"{XlsxConstants.WorksheetRelType}\" Target=\"worksheets/sheet{sheet.SheetId}.bin\"/>");
            }
            sb.Append("</Relationships>");
            return sb.ToString();
        }

        private void WriteWorkbookRels()
        {
            WriteEntry("xl/_rels/workbook.bin.rels", BuildWorkbookRelsXml());
        }

        private ValueTask WriteWorkbookRelsAsync(CancellationToken ct)
        {
            return WriteEntryAsync("xl/_rels/workbook.bin.rels", BuildWorkbookRelsXml(), ct);
        }

        private void BuildStylesBin(BiffBuffer payload, BiffBuffer data)
        {
            Dictionary<string, int> numFmtIds = _styles.AssignCustomNumberFormatIds();
            Biff12RecordWriter.WriteRecord(data, Brt.BeginStyleSheet);
            WriteCountedRecord(data, payload, Brt.BeginFmts, numFmtIds.Count);
            foreach (var (format, id) in numFmtIds)
            {
                WriteFmt(data, payload, id, format);
            }
            Biff12RecordWriter.WriteRecord(data, Brt.EndFmts);
            WriteCountedRecord(data, payload, Brt.BeginFonts, 1);
            WriteBlobRecord(data, payload, Brt.Font, DefaultFontPayload);
            Biff12RecordWriter.WriteRecord(data, Brt.EndFonts);
            WriteCountedRecord(data, payload, Brt.BeginFills, 2);
            WriteFill(data, payload, FillPatternNone);
            WriteFill(data, payload, FillPatternGray125);
            Biff12RecordWriter.WriteRecord(data, Brt.EndFills);
            WriteCountedRecord(data, payload, Brt.BeginBorders, 1);
            WriteBlobRecord(data, payload, Brt.Border, DefaultBorderPayload);
            Biff12RecordWriter.WriteRecord(data, Brt.EndBorders);
            WriteCountedRecord(data, payload, Brt.BeginCellStyleXFs, 1);
            WriteXf(data, payload, 0, isStyleXf: true);
            Biff12RecordWriter.WriteRecord(data, Brt.EndCellStyleXFs);
            IReadOnlyList<CellStyle> styles = _styles.Styles;
            WriteCountedRecord(data, payload, Brt.BeginCellXFs, styles.Count);
            WriteXf(data, payload, 0);
            WriteXf(data, payload, 14);
            for (int i = 2; i < styles.Count; i++)
            {
                int numFmtId = styles[i].NumberFormat is string format ? numFmtIds[format] : 0;
                WriteXf(data, payload, numFmtId);
            }
            Biff12RecordWriter.WriteRecord(data, Brt.EndCellXFs);
            WriteStyles(data, payload);
            Biff12RecordWriter.WriteRecord(data, Brt.EndStyleSheet);
        }

        private void WriteStyles()
        {
            using BiffBuffer payload = new(96);
            using BiffBuffer data = new(512);
            BuildStylesBin(payload, data);
            WriteEntry("xl/styles.bin", data.Span);
        }

        private async ValueTask WriteStylesAsync(CancellationToken ct)
        {
            using BiffBuffer payload = new(96);
            using BiffBuffer data = new(512);
            BuildStylesBin(payload, data);
            await WriteEntryAsync("xl/styles.bin", data.Memory, ct).ConfigureAwait(false);
        }

        private static void WriteStyles(BiffBuffer data, BiffBuffer payload)
        {
            WriteCountedRecord(data, payload, Brt.BeginStyles, 1);
            payload.Reset();
            payload.WriteU32(0);
            payload.WriteU16(StyleBuiltIn);
            payload.WriteByte(0);
            payload.WriteByte(0);
            Biff12RecordWriter.WriteWideString(payload, "Normal");
            Biff12RecordWriter.WriteRecord(data, Brt.Style, payload.Span);
            Biff12RecordWriter.WriteRecord(data, Brt.EndStyles);
        }

        private const int StyleBuiltIn = 0x0001;

        private static void WriteFmt(BiffBuffer data, BiffBuffer payload, int numFmtId, string formatCode)
        {
            payload.Reset();
            payload.WriteU16(numFmtId);
            Biff12RecordWriter.WriteWideString(payload, formatCode);
            Biff12RecordWriter.WriteRecord(data, Brt.Fmt, payload.Span);
        }

        private void WriteSharedStrings()
        {
            using var data = _sharedStrings is null ? EmptySharedStrings() : _sharedStrings.ToXlsbBytes();
            WriteEntry("xl/sharedStrings.bin", data.Span);
        }

        private async ValueTask WriteSharedStringsAsync(CancellationToken ct)
        {
            using var data = _sharedStrings is null ? EmptySharedStrings() : _sharedStrings.ToXlsbBytes();
            await WriteEntryAsync("xl/sharedStrings.bin", data.Memory, ct).ConfigureAwait(false);
        }

        private static BiffBuffer EmptySharedStrings()
        {
            var data = new BiffBuffer(16);
            using var payload = new BiffBuffer(8);
            payload.WriteU32(0);
            payload.WriteU32(0);
            Biff12RecordWriter.WriteRecord(data, Brt.BeginSst, payload.Span);
            Biff12RecordWriter.WriteRecord(data, Brt.EndSst);
            return data;
        }

        private static void WriteCountedRecord(BiffBuffer data, BiffBuffer payload, int id, int count)
        {
            payload.Reset();
            payload.WriteU32((uint)count);
            Biff12RecordWriter.WriteRecord(data, id, payload.Span);
        }

        private static void WriteBlobRecord(BiffBuffer data, BiffBuffer payload, int id, ReadOnlySpan<byte> blob)
        {
            payload.Reset();
            payload.Write(blob);
            Biff12RecordWriter.WriteRecord(data, id, payload.Span);
        }

        private static ReadOnlySpan<byte> DefaultFontPayload => [
            0xDC, 0x00, 0x00, 0x00, 0x90, 0x01, 0x00, 0x00,
            0x00, 0x02, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
            0x00, 0x00, 0x00, 0x00, 0x00, 0x07, 0x00, 0x00,
            0x00, 0x43, 0x00, 0x61, 0x00, 0x6C, 0x00, 0x69,
            0x00, 0x62, 0x00, 0x72, 0x00, 0x69, 0x00,
        ];
        private static ReadOnlySpan<byte> FillPayloadAfterPattern => [
            0x03, 0x40, 0x00, 0x00, 0x00, 0x00, 0x00, 0xFF,
            0x03, 0x41, 0x00, 0x00, 0xFF, 0xFF, 0xFF, 0xFF,
            0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
            0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
            0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
            0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
            0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
            0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
        ];

        private const uint FillPatternNone = 0x00;
        private const uint FillPatternGray125 = 0x11;

        private static void WriteFill(BiffBuffer data, BiffBuffer payload, uint pattern)
        {
            payload.Reset();
            payload.WriteU32(pattern);
            payload.Write(FillPayloadAfterPattern);
            Biff12RecordWriter.WriteRecord(data, Brt.Fill, payload.Span);
        }
        private static ReadOnlySpan<byte> DefaultBorderPayload => [
            0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x00,
            0x00, 0x00, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00,
            0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x01,
            0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
            0x00, 0x01, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
            0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x00,
            0x00, 0x00, 0x00,
        ];

        private const string AppPropertiesXml = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
            "<Properties xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/extended-properties\">" +
            "<Application>ExcelReader</Application><AppVersion>1.0000</AppVersion></Properties>";

        private void WriteAppProperties()
        {
            WriteEntry("docProps/app.xml", AppPropertiesXml);
        }

        private ValueTask WriteAppPropertiesAsync(CancellationToken ct)
        {
            return WriteEntryAsync("docProps/app.xml", AppPropertiesXml, ct);
        }

        private static void WriteXf(BiffBuffer data, BiffBuffer payload, int numFmtId, bool isStyleXf = false)
        {
            payload.Reset();
            payload.WriteU16(isStyleXf ? ushort.MaxValue : 0);
            payload.WriteU16(numFmtId);
            payload.WriteU16(0);
            payload.WriteU16(0);
            payload.WriteU16(0);
            payload.WriteByte(0);
            payload.WriteByte(0);
            payload.WriteU16(XfDefaultFlags);
            payload.WriteU16(numFmtId != 0 ? XfAttributeNumberFormat : 0);
            Biff12RecordWriter.WriteRecord(data, Brt.Xf, payload.Span);
        }

        private const int XfDefaultFlags = 0x1010;
        private const int XfAttributeNumberFormat = 0x0001;

        private string BuildContentTypesXml()
        {
            StringBuilder sb = new();
            sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
            sb.Append(CultureInfo.InvariantCulture, $"<Types xmlns=\"{XlsxConstants.ContentTypesNs}\">");
            sb.Append(CultureInfo.InvariantCulture, $"<Default Extension=\"rels\" ContentType=\"{XlsxConstants.RelationshipsContentType}\"/>");
            sb.Append("<Default Extension=\"xml\" ContentType=\"application/xml\"/>");
            sb.Append("<Default Extension=\"bin\" ContentType=\"application/vnd.ms-excel.sheet.binary.macroEnabled.main\"/>");
            sb.Append("<Override PartName=\"/xl/workbook.bin\" ContentType=\"application/vnd.ms-excel.sheet.binary.macroEnabled.main\"/>");
            sb.Append("<Override PartName=\"/xl/styles.bin\" ContentType=\"application/vnd.ms-excel.styles\"/>");
            sb.Append("<Override PartName=\"/xl/sharedStrings.bin\" ContentType=\"application/vnd.ms-excel.sharedStrings\"/>");
            sb.Append("<Override PartName=\"/docProps/app.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.extended-properties+xml\"/>");
            foreach (ref readonly var sheet in CollectionsMarshal.AsSpan(_sheets))
            {
                sb.Append(CultureInfo.InvariantCulture, $"<Override PartName=\"/xl/worksheets/sheet{sheet.SheetId}.bin\" ContentType=\"application/vnd.ms-excel.worksheet\"/>");
            }
            sb.Append("</Types>");
            return sb.ToString();
        }

        private void WriteContentTypes()
        {
            WriteEntry("[Content_Types].xml", BuildContentTypesXml());
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
    }
}
