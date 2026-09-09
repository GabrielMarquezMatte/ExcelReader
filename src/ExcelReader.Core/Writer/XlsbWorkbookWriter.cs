using System.Globalization;
using System.IO.Compression;
using System.Runtime.InteropServices;
using System.Text;
using ExcelReader.Core.Internal;
using ExcelReader.Core.Reader;
using ExcelReader.Core.Writer.Internal;

namespace ExcelReader.Core.Writer
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
        private WriterState _state = WriterState.Created;
        private XlsbSheetWriter? _activeSheet;
        private bool _disposed;

        private XlsbWorkbookWriter(ZipArchive zip, Stream stream, bool leaveOpen, bool date1904, bool useSharedStrings, CompressionLevel compression, bool prefetchWrite)
        {
            _zip = zip;
            _stream = stream;
            _leaveOpen = leaveOpen;
            _date1904 = date1904;
            UseSharedStrings = useSharedStrings;
            _compression = compression;
            _prefetchWrite = prefetchWrite;
            _sharedStrings = useSharedStrings ? new SharedStringTable() : null;
        }

        /// <summary>Creates a writer that will produce an .xlsb archive on <paramref name="stream"/> once started.</summary>
        /// <param name="stream">The destination stream; must be writable.</param>
        /// <param name="leaveOpen">If <see langword="true"/>, <paramref name="stream"/> is not disposed when the writer is disposed.</param>
        /// <param name="date1904">Whether the workbook uses the 1904 date system instead of the default 1900 system.</param>
        /// <param name="compression">The zip compression level to use for every part written.</param>
        /// <param name="useSharedStrings">Whether string cells are deduplicated through a shared string table instead of written inline.</param>
        /// <param name="prefetchWrite">
        /// When <see langword="true"/>, each sheet's deflate runs on a background thread instead of the
        /// calling thread, overlapping compression with row serialization. Defaults to <see langword="false"/>.
        /// Mirrors <see cref="ExcelReaderOptions.PrefetchDecompression"/>'s tradeoff on the read
        /// side: worth it for single-file batch writing, not for a concurrent server workload already
        /// saturating the CPU.
        /// </param>
        /// <param name="ct">A token to cancel creation before any I/O has started.</param>
        public static ValueTask<XlsbWorkbookWriter> CreateAsync(
            Stream stream,
            bool leaveOpen = false,
            bool date1904 = false,
            CompressionLevel compression = CompressionLevel.Fastest,
            bool useSharedStrings = false,
            bool prefetchWrite = false,
            CancellationToken ct = default)
        {
            ArgumentNullException.ThrowIfNull(stream);
            ct.ThrowIfCancellationRequested();
            ZipArchive zip = new(stream, ZipArchiveMode.Create, leaveOpen: true);
            return ValueTask.FromResult(new XlsbWorkbookWriter(zip, stream, leaveOpen, date1904, useSharedStrings, compression, prefetchWrite));
        }

        /// <summary>
        /// Synchronous counterpart to <see cref="CreateAsync"/>, for native/unmanaged callers whose ABI
        /// is synchronous. Parameters mirror <see cref="CreateAsync"/> exactly, minus <c>ct</c>.
        /// </summary>
        public static XlsbWorkbookWriter Create(
            Stream stream,
            bool leaveOpen = false,
            bool date1904 = false,
            CompressionLevel compression = CompressionLevel.Fastest,
            bool useSharedStrings = false,
            bool prefetchWrite = false)
        {
            ArgumentNullException.ThrowIfNull(stream);
            ZipArchive zip = new(stream, ZipArchiveMode.Create, leaveOpen: true);
            return new XlsbWorkbookWriter(zip, stream, leaveOpen, date1904, useSharedStrings, compression, prefetchWrite);
        }

        /// <summary>
        /// Synchronous counterpart to <see cref="StartAsync"/>, for native/unmanaged callers whose ABI
        /// is synchronous.
        /// </summary>
        public void Start()
        {
            WriterStateGuard.ThrowIfEnded(_state, this);
            WriterStateGuard.RequireCreated(_state, nameof(XlsbWorkbookWriter));
            _state = WriterState.Started;
        }

        /// <inheritdoc/>
        public ValueTask StartAsync(CancellationToken ct = default)
        {
            WriterStateGuard.ThrowIfEnded(_state, this);
            WriterStateGuard.RequireCreated(_state, nameof(XlsbWorkbookWriter));
            ct.ThrowIfCancellationRequested();
            _state = WriterState.Started;
            return ValueTask.CompletedTask;
        }

        /// <inheritdoc/>
        public XlsbSheetWriter AddSheet(string name)
        {
            WriterStateGuard.RequireCanAddSheet(
                _state, this, nameof(XlsbWorkbookWriter), name, _activeSheet is not null, nameof(XlsbSheetWriter));
            int sheetId = _sheets.Count + 1;
            _activeSheet = new XlsbSheetWriter(this, _zip, name, sheetId, _date1904, _compression, _prefetchWrite);
            return _activeSheet;
        }

        internal void RegisterSheet(XlsbSheetWriter sheet)
        {
            _sheets.Add(sheet);
        }

        internal void NotifySheetEnded()
        {
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
        public void End()
        {
            WriterStateGuard.ThrowIfEnded(_state, this);
            WriterStateGuard.RequireStarted(_state, nameof(XlsbWorkbookWriter), "ending");
            // Must dispose (and thus register) the active sheet before checking _sheets.Count.
            _activeSheet?.Dispose();
            if (_sheets.Count == 0)
            {
                throw new InvalidOperationException("A workbook must contain at least one sheet.");
            }
            _state = WriterState.Ended;

            WriteRootRels();
            WriteWorkbook();
            WriteWorkbookRels();
            WriteStyles();
            WriteSharedStrings();
            WriteAppProperties();
            WriteContentTypes();
            ZipArchiveDisposal.Dispose(_zip);
        }

        /// <inheritdoc/>
        public async ValueTask EndAsync(CancellationToken ct = default)
        {
            WriterStateGuard.ThrowIfEnded(_state, this);
            WriterStateGuard.RequireStarted(_state, nameof(XlsbWorkbookWriter), "ending");
            ct.ThrowIfCancellationRequested();
            // XlsbSheetWriter registers itself in EndAsync/DisposeAsync, not StartAsync, so the active
            // sheet must be disposed (and thus registered) before checking _sheets.Count below.
            if (_activeSheet is not null)
            {
                await _activeSheet.DisposeAsync().ConfigureAwait(false);
            }
            if (_sheets.Count == 0)
            {
                throw new InvalidOperationException("A workbook must contain at least one sheet.");
            }
            _state = WriterState.Ended;

            await WriteRootRelsAsync(ct).ConfigureAwait(false);
            await WriteWorkbookAsync(ct).ConfigureAwait(false);
            await WriteWorkbookRelsAsync(ct).ConfigureAwait(false);
            await WriteStylesAsync(ct).ConfigureAwait(false);
            await WriteSharedStringsAsync(ct).ConfigureAwait(false);
            await WriteAppPropertiesAsync(ct).ConfigureAwait(false);
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
        public void Dispose()
        {
            if (_disposed)
            {
                return;
            }
            _disposed = true;
            if (_state == WriterState.Started)
            {
                if (_sheets.Count == 0 && _activeSheet is null)
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
                // _activeSheet is checked too: XlsbSheetWriter only registers itself in EndAsync, so a
                // still-open sheet must route through EndAsync rather than being silently abandoned.
                if (_sheets.Count == 0 && _activeSheet is null)
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
            foreach (ref readonly var sheet in CollectionsMarshal.AsSpan(_sheets))
            {
                payload.Reset();
                payload.WriteU32(0);
                payload.WriteU32((uint)sheet.SheetId);
                Biff12RecordWriter.WriteWideString(payload, $"s{sheet.SheetId}");
                Biff12RecordWriter.WriteWideString(payload, sheet.Name);
                Biff12RecordWriter.WriteRecord(data, Brt.BundleSh, payload.Span);
            }
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

        // Bold/italic not represented here: BrtFont's payload is an opaque, byte-exact blob
        // (DefaultFontPayload) reverse-engineered from a real file, and no field in it is safe to flip
        // without a verified [MS-XLSB] field map. Every custom style's Xf keeps font index 0.
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

        // Bold/italic not represented here: BrtFont's payload is an opaque, byte-exact blob
        // (DefaultFontPayload) reverse-engineered from a real file, and no field in it is safe to flip
        // without a verified [MS-XLSB] field map. Every custom style's Xf keeps font index 0.
        private async ValueTask WriteStylesAsync(CancellationToken ct)
        {
            using BiffBuffer payload = new(96);
            using BiffBuffer data = new(512);
            BuildStylesBin(payload, data);
            await WriteEntryAsync("xl/styles.bin", data.Memory, ct).ConfigureAwait(false);
        }

        // The STYLES production is mandatory in a styles part; one built-in "Normal" cell style
        // pointing at cellStyleXfs[0] is the minimum Excel accepts without repairing the workbook.
        //
        // BrtStyle payload, verified byte-for-byte against a real Excel-authored .xlsb:
        //   ixf (u32) | grbit (u16) | iStyBuiltIn (u8) | iLevel (u8) | stName (XLWideString)
        private static void WriteStyles(BiffBuffer data, BiffBuffer payload)
        {
            WriteCountedRecord(data, payload, Brt.BeginStyles, 1);
            payload.Reset();
            payload.WriteU32(0);            // ixf -> cellStyleXfs[0]
            payload.WriteU16(StyleBuiltIn); // grbit
            payload.WriteByte(0);           // iStyBuiltIn: 0 == the "Normal" built-in style
            payload.WriteByte(0);           // iLevel
            Biff12RecordWriter.WriteWideString(payload, "Normal");
            Biff12RecordWriter.WriteRecord(data, Brt.Style, payload.Span);
            Biff12RecordWriter.WriteRecord(data, Brt.EndStyles);
        }

        // BrtStyle grbit bit 0: the style is one of Excel's built-ins rather than a user-defined one.
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

        // Fixed byte blobs Excel expects verbatim; nothing in them varies per workbook.
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
        // BrtFill is fls (u32 pattern type) followed by this 64-byte remainder — the fg/bg BrtColor
        // pair plus the gradient-stop area a solid fill leaves zeroed. Verified byte-for-byte against
        // a real Excel-authored .xlsb.
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

        // Excel requires the first two fills to be exactly these, in this order, in every styles part.
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

        // BrtXF, 16 bytes ([MS-XLSB] 2.4.816). Field-by-field, not a partial blob: every field must be
        // explicitly zeroed rather than left to whatever garbage a stackalloc might carry, since a
        // stray value here produces an index pointing at a font/fill the part never declared.
        private static void WriteXf(BiffBuffer data, BiffBuffer payload, int numFmtId, bool isStyleXf = false)
        {
            payload.Reset();
            payload.WriteU16(isStyleXf ? ushort.MaxValue : 0); // ixfeParent (0xFFFF for a style XF)
            payload.WriteU16(numFmtId);                        // iFmt
            payload.WriteU16(0);                               // iFont
            payload.WriteU16(0);                               // iFill
            payload.WriteU16(0);                               // ixBorder
            payload.WriteByte(0);                              // trot
            payload.WriteByte(0);                              // indent
            payload.WriteU16(XfDefaultFlags);
            payload.WriteU16(numFmtId != 0 ? XfAttributeNumberFormat : 0); // xfGrbitAtr
            Biff12RecordWriter.WriteRecord(data, Brt.Xf, payload.Span);
        }

        // Vertical alignment "bottom" plus fLocked, Excel's default cell protection state.
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
