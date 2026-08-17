using System.Diagnostics.CodeAnalysis;
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
        [SuppressMessage("IDisposableAnalyzers.Correctness", "IDISP001:Dispose created",
            Justification = "Factory method transfers ZipArchive ownership to XlsbWorkbookWriter; caller disposes via DisposeAsync.")]
        [SuppressMessage("IDisposableAnalyzers.Correctness", "IDISP004:Don't ignore created IDisposable",
            Justification = "Factory method transfers ZipArchive ownership to XlsbWorkbookWriter; caller disposes via DisposeAsync.")]
        [SuppressMessage("Reliability", "CA2000:Dispose objects before losing scope",
            Justification = "XlsbWorkbookWriter takes ownership of ZipArchive and disposes it in DisposeAsync/EndAsync.")]
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
        [SuppressMessage("IDisposableAnalyzers.Correctness", "IDISP001:Dispose created",
            Justification = "Factory method transfers ZipArchive ownership to XlsbWorkbookWriter; caller disposes via Dispose/DisposeAsync.")]
        [SuppressMessage("IDisposableAnalyzers.Correctness", "IDISP004:Don't ignore created IDisposable",
            Justification = "Factory method transfers ZipArchive ownership to XlsbWorkbookWriter; caller disposes via Dispose/DisposeAsync.")]
        [SuppressMessage("Reliability", "CA2000:Dispose objects before losing scope",
            Justification = "XlsbWorkbookWriter takes ownership of ZipArchive and disposes it in Dispose(Async)/End(Async).")]
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
        [SuppressMessage("IDisposableAnalyzers.Correctness", "IDISP003:Dispose previous before re-assigning",
            Justification = "RequireCanAddSheet above guarantees _activeSheet is null (any previous sheet was already ended and disposed) before this assigns a new one.")]
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

        [SuppressMessage("IDisposableAnalyzers.Correctness", "IDISP003:Dispose previous before re-assigning",
            Justification = "Called only after the active sheet's own End/Dispose has already run (from End/Dispose below); this just clears the now-stale reference.")]
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

        // Valid styleId range for SetColumnStyle/StartRowAsync(int,...) is [0, StyleCount).
        internal int StyleCount => _styles.Count;

        /// <summary>
        /// Synchronous counterpart to <see cref="EndAsync"/>, for native/unmanaged callers whose ABI is
        /// synchronous.
        /// </summary>
        public void End()
        {
            WriterStateGuard.ThrowIfEnded(_state, this);
            WriterStateGuard.RequireStarted(_state, nameof(XlsbWorkbookWriter), "ending");
            // See EndAsync's remarks: the active sheet must be disposed (and thus registered) before
            // checking _sheets.Count.
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
            // Unlike XlsxSheetWriter (which registers itself in StartAsync, so _sheets.Count is
            // already accurate by the time the workbook ends), XlsbSheetWriter registers itself in its
            // own EndAsync/DisposeAsync — so the active sheet must be disposed (and thus registered)
            // BEFORE checking _sheets.Count here, or a workbook with exactly one still-open sheet (never
            // explicitly ended by the caller, relying on this method to do it) would wrongly see zero
            // sheets and reject a perfectly valid workbook. The zero-sheet check must still run before
            // _state flips to Ended, so a genuinely empty workbook leaves DisposeAsync able to find and
            // release _zip.
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
                // End deliberately rejects a zero-sheet workbook — see DisposeAsync's remarks.
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
                // EndAsync deliberately rejects a zero-sheet workbook, so disposal
                // must release a partially configured writer itself rather than routing through it —
                // this branch used to just flip _state without disposing _zip at all, unlike
                // XlsxWorkbookWriter's equivalent branch. _activeSheet is checked too (unlike Xlsx's
                // equivalent) because XlsbSheetWriter only registers itself in EndAsync, not StartAsync
                // — so a still-open sheet here must route through EndAsync to register (and properly
                // end) it, rather than being silently abandoned by the "truly nothing to do" branch.
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

        // Bold/italic are deliberately not represented here: BrtFont's payload is an opaque, byte-exact
        // blob (DefaultFontPayload below) reverse-engineered from a real file, and no field in it is
        // safe to flip without a verified [MS-XLSB] field map. Every custom style's Xf keeps font index
        // 0; only its number format varies. XLSX gets full Bold/Italic support because its font XML
        // element is self-describing (<b/>/<i/>) instead of an opaque binary blob.
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

        // Bold/italic are deliberately not represented here: BrtFont's payload is an opaque, byte-exact
        // blob (DefaultFontPayload below) reverse-engineered from a real file, and no field in it is
        // safe to flip without a verified [MS-XLSB] field map. Every custom style's Xf keeps font index
        // 0; only its number format varies. XLSX gets full Bold/Italic support because its font XML
        // element is self-describing (<b/>/<i/>) instead of an opaque binary blob.
        private async ValueTask WriteStylesAsync(CancellationToken ct)
        {
            using BiffBuffer payload = new(96);
            using BiffBuffer data = new(512);
            BuildStylesBin(payload, data);
            await WriteEntryAsync("xl/styles.bin", data.Memory, ct).ConfigureAwait(false);
        }

        // The STYLES production is mandatory in a styles part — omitting it makes Excel report
        // "Formato de parte de /xl/styles.bin" and silently repair the workbook on open. One built-in
        // "Normal" cell style pointing at cellStyleXfs[0] is the minimum Excel accepts; the optional
        // DXFS/TABLESTYLES/slicer blocks a full Excel-authored file also carries stay omitted.
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

        // The default font/fill/border records are fixed byte blobs Excel expects verbatim in a
        // minimal styles part; nothing in them varies per workbook, so they are emitted from constants
        // rather than composed field by field.
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
        // a real Excel-authored .xlsb, whose two fills differ from each other in fls alone.
        //
        // This blob used to be written with its leading fls byte missing, shifting every field left by
        // one so fls decoded as 0x03000000 instead of 0 — enough for Excel to reject the fills block
        // and repair the workbook ("Formato de parte de /xl/styles.bin").
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

        // Excel requires the first two fills to be exactly these, in this order, in every styles part —
        // the same fixed pair its own XLSX output always carries. A workbook with fewer is repaired.
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

        // BrtXF, 16 bytes ([MS-XLSB] 2.4.816). Field-by-field rather than a partial blob: this used to
        // write its middle 10 bytes from `stackalloc byte[10]` under [SkipLocalsInit], which does NOT
        // zero the allocation — so iFont/iFill/ixBorder/trot/indent/flags were filled with whatever
        // happened to be on the stack. That produced indices pointing at fonts and fills the part never
        // declared, which Excel repairs ("Formato de parte de /xl/styles.bin"). The values below are
        // taken from a real Excel-authored .xlsb, whose style and cell XFs differ only in ixfeParent.
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
            // xfGrbitAtr marks which attributes this XF overrides from its parent style; a custom
            // number format is the only one this writer ever sets.
            payload.WriteU16(numFmtId != 0 ? XfAttributeNumberFormat : 0);
            Biff12RecordWriter.WriteRecord(data, Brt.Xf, payload.Span);
        }

        // Alignment/protection bits every XF in an Excel-written styles part carries: vertical
        // alignment "bottom" plus fLocked, Excel's default cell protection state.
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
