using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.IO.Compression;
using System.Runtime.InteropServices;
using System.Text;
using ExcelReader.Core.Reader;
using ExcelReader.Core.Writer.Internal;

namespace ExcelReader.Core.Writer
{
    public sealed class XlsbWorkbookWriter : IWorkbookWriter<XlsbSheetWriter>
    {
        private const int MaxSheetNameLength = 31;

        private readonly ZipArchive _zip;
        private readonly Stream _stream;
        private readonly bool _leaveOpen;
        private readonly bool _date1904;
        private readonly CompressionLevel _compression;
        private readonly SharedStringTable? _sharedStrings;
        private readonly List<XlsbSheetWriter> _sheets = [];
        private WriterState _state = WriterState.Created;
        private XlsbSheetWriter? _activeSheet;
        private bool _disposed;

        private XlsbWorkbookWriter(ZipArchive zip, Stream stream, bool leaveOpen, bool date1904, bool useSharedStrings, CompressionLevel compression)
        {
            _zip = zip;
            _stream = stream;
            _leaveOpen = leaveOpen;
            _date1904 = date1904;
            UseSharedStrings = useSharedStrings;
            _compression = compression;
            _sharedStrings = useSharedStrings ? new SharedStringTable() : null;
        }

        [SuppressMessage("IDisposableAnalyzers.Correctness", "IDISP001:Dispose created",
            Justification = "Factory method transfers ZipArchive ownership to XlsbWorkbookWriter; caller disposes via DisposeAsync.")]
        [SuppressMessage("Reliability", "CA2000:Dispose objects before losing scope",
            Justification = "XlsbWorkbookWriter takes ownership of ZipArchive and disposes it in DisposeAsync/EndAsync.")]
        [SuppressMessage("Design", "CA1068:CancellationToken parameters must come last",
            Justification = "useSharedStrings was added after existing optional parameters to preserve positional source compatibility.")]
        public static ValueTask<XlsbWorkbookWriter> CreateAsync(
            Stream stream,
            bool leaveOpen = false,
            bool date1904 = false,
            CompressionLevel compression = CompressionLevel.Fastest,
            CancellationToken ct = default,
            bool useSharedStrings = false)
        {
            ArgumentNullException.ThrowIfNull(stream);
            ct.ThrowIfCancellationRequested();
            ZipArchive zip = new(stream, ZipArchiveMode.Create, leaveOpen: true);
            return ValueTask.FromResult(new XlsbWorkbookWriter(zip, stream, leaveOpen, date1904, useSharedStrings, compression));
        }

        public ValueTask StartAsync(CancellationToken ct = default)
        {
            ObjectDisposedException.ThrowIf(_state == WriterState.Ended, this);
            if (_state != WriterState.Created)
            {
                throw new InvalidOperationException("XlsbWorkbookWriter has already been started.");
            }
            ct.ThrowIfCancellationRequested();
            _state = WriterState.Started;
            return ValueTask.CompletedTask;
        }

        public XlsbSheetWriter AddSheet(string name)
        {
            ArgumentNullException.ThrowIfNull(name);
            ObjectDisposedException.ThrowIf(_state == WriterState.Ended, this);
            if (_state != WriterState.Started)
            {
                throw new InvalidOperationException("XlsbWorkbookWriter must be started before adding sheets.");
            }
            if (name.Length is 0 or > MaxSheetNameLength)
            {
                throw new ArgumentException($"Sheet names must be 1 to {MaxSheetNameLength} characters.", nameof(name));
            }
            if (_activeSheet is not null)
            {
                throw new InvalidOperationException("The previous XlsbSheetWriter must be ended before adding a new sheet.");
            }
            int sheetId = _sheets.Count + 1;
            _activeSheet = new XlsbSheetWriter(this, _zip, name, sheetId, _date1904, _compression);
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

        public async ValueTask EndAsync(CancellationToken ct = default)
        {
            ObjectDisposedException.ThrowIf(_state == WriterState.Ended, this);
            if (_state != WriterState.Started)
            {
                throw new InvalidOperationException("XlsbWorkbookWriter must be started before ending.");
            }
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

            await WriteRootRelsAsync(ct).ConfigureAwait(false);
            await WriteWorkbookAsync(ct).ConfigureAwait(false);
            await WriteWorkbookRelsAsync(ct).ConfigureAwait(false);
            await WriteStylesAsync(ct).ConfigureAwait(false);
            await WriteSharedStringsAsync(ct).ConfigureAwait(false);
            await WriteAppPropertiesAsync(ct).ConfigureAwait(false);
            await WriteContentTypesAsync(ct).ConfigureAwait(false);
#if NET10_0_OR_GREATER
            await _zip.DisposeAsync().ConfigureAwait(false);
#else
            _zip.Dispose();
#endif
        }

        public ValueTask FlushAsync(CancellationToken ct = default)
        {
            return ZipEntryWriter.FlushAsync(_stream, ct);
        }

        public async ValueTask DisposeAsync()
        {
            if (_disposed)
            {
                return;
            }
            _disposed = true;
            if (_state == WriterState.Started)
            {
                if (_sheets.Count > 0 || _activeSheet is not null)
                {
                    await EndAsync().ConfigureAwait(false);
                }
                else
                {
                    _state = WriterState.Ended;
                }
            }
            else if (_state == WriterState.Created)
            {
#if NET10_0_OR_GREATER
                await _zip.DisposeAsync().ConfigureAwait(false);
#else
                _zip.Dispose();
#endif
            }
            if (!_leaveOpen)
            {
                await _stream.DisposeAsync().ConfigureAwait(false);
            }
        }

        private ValueTask WriteRootRelsAsync(CancellationToken ct)
        {
            const string xml = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                $"<Relationships xmlns=\"{XlsxConstants.PackageRelationshipsNs}\">" +
                "<Relationship Id=\"app\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties\" Target=\"docProps/app.xml\"/>" +
                $"<Relationship Id=\"wb\" Type=\"{XlsxConstants.WorkbookRelType}\" Target=\"xl/workbook.bin\"/>" +
                "</Relationships>";
            return WriteEntryAsync("_rels/.rels", xml, ct);
        }

        private async ValueTask WriteWorkbookAsync(CancellationToken ct)
        {
            using BiffBuffer payload = new(256);
            using BiffBuffer data = new(1024);
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
            await WriteEntryAsync("xl/workbook.bin", data.Memory, ct).ConfigureAwait(false);
        }

        private ValueTask WriteWorkbookRelsAsync(CancellationToken ct)
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
            return WriteEntryAsync("xl/_rels/workbook.bin.rels", sb.ToString(), ct);
        }

        private async ValueTask WriteStylesAsync(CancellationToken ct)
        {
            using BiffBuffer payload = new(96);
            using BiffBuffer data = new(512);
            Biff12RecordWriter.WriteRecord(data, Brt.BeginStyleSheet);
            WriteCountedRecord(data, payload, Brt.BeginFmts, 0);
            Biff12RecordWriter.WriteRecord(data, Brt.EndFmts);
            WriteCountedRecord(data, payload, Brt.BeginFonts, 1);
            WriteDefaultFont(data, payload);
            Biff12RecordWriter.WriteRecord(data, Brt.EndFonts);
            WriteCountedRecord(data, payload, Brt.BeginFills, 1);
            WriteDefaultFill(data, payload);
            Biff12RecordWriter.WriteRecord(data, Brt.EndFills);
            WriteCountedRecord(data, payload, Brt.BeginBorders, 1);
            WriteDefaultBorder(data, payload);
            Biff12RecordWriter.WriteRecord(data, Brt.EndBorders);
            WriteCountedRecord(data, payload, Brt.BeginCellStyleXFs, 1);
            WriteXf(data, payload, 0, isStyleXf: true);
            Biff12RecordWriter.WriteRecord(data, Brt.EndCellStyleXFs);
            WriteCountedRecord(data, payload, Brt.BeginCellXFs, 2);
            WriteXf(data, payload, 0);
            WriteXf(data, payload, 14);
            Biff12RecordWriter.WriteRecord(data, Brt.EndCellXFs);
            Biff12RecordWriter.WriteRecord(data, Brt.EndStyleSheet);
            await WriteEntryAsync("xl/styles.bin", data.Memory, ct).ConfigureAwait(false);
        }

        private ValueTask WriteSharedStringsAsync(CancellationToken ct)
        {
            ReadOnlyMemory<byte> data = _sharedStrings is null
                ? EmptySharedStrings()
                : _sharedStrings.ToXlsbBytes();
            return WriteEntryAsync("xl/sharedStrings.bin", data, ct);
        }

        private static ReadOnlyMemory<byte> EmptySharedStrings()
        {
            using var data = new BiffBuffer(16);
            using var payload = new BiffBuffer(8);
            payload.WriteU32(0);
            payload.WriteU32(0);
            Biff12RecordWriter.WriteRecord(data, Brt.BeginSst, payload.Span);
            Biff12RecordWriter.WriteRecord(data, Brt.EndSst);
            return data.Memory.ToArray();
        }

        private static void WriteCountedRecord(BiffBuffer data, BiffBuffer payload, int id, int count)
        {
            payload.Reset();
            payload.WriteU32((uint)count);
            Biff12RecordWriter.WriteRecord(data, id, payload.Span);
        }
        private static ReadOnlySpan<byte> DefaultFontPayload => [
            0xDC, 0x00, 0x00, 0x00, 0x90, 0x01, 0x00, 0x00,
            0x00, 0x02, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
            0x00, 0x00, 0x00, 0x00, 0x00, 0x07, 0x00, 0x00,
            0x00, 0x43, 0x00, 0x61, 0x00, 0x6C, 0x00, 0x69,
            0x00, 0x62, 0x00, 0x72, 0x00, 0x69, 0x00,
        ];
        private static void WriteDefaultFont(BiffBuffer data, BiffBuffer payload)
        {
            payload.Reset();
            payload.Write(DefaultFontPayload);
            Biff12RecordWriter.WriteRecord(data, 43, payload.Span);
        }
        private static ReadOnlySpan<byte> DefaultFillPayload => [
            0x00, 0x00, 0x00, 0x03, 0x40, 0x00, 0x00, 0x00,
            0x00, 0x00, 0xFF, 0x03, 0x41, 0x00, 0x00, 0xFF,
            0xFF, 0xFF, 0xFF, 0x00, 0x00, 0x00, 0x00, 0x00,
            0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
            0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
            0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
            0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
            0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
            0x00, 0x00, 0x00, 0x00,
        ];
        private static void WriteDefaultFill(BiffBuffer data, BiffBuffer payload)
        {
            payload.Reset();
            payload.Write(DefaultFillPayload);
            Biff12RecordWriter.WriteRecord(data, 45, payload.Span);
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
        private static void WriteDefaultBorder(BiffBuffer data, BiffBuffer payload)
        {
            payload.Reset();
            payload.Write(DefaultBorderPayload);
            Biff12RecordWriter.WriteRecord(data, 46, payload.Span);
        }

        private ValueTask WriteAppPropertiesAsync(CancellationToken ct)
        {
            const string xml = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                "<Properties xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/extended-properties\">" +
                "<Application>ExcelReader</Application><AppVersion>1.0000</AppVersion></Properties>";
            return WriteEntryAsync("docProps/app.xml", xml, ct);
        }

        private static void WriteXf(BiffBuffer data, BiffBuffer payload, int numFmtId, bool isStyleXf = false)
        {
            payload.Reset();
            payload.WriteU16(isStyleXf ? ushort.MaxValue : (ushort)0);
            payload.WriteU16(numFmtId);
            ReadOnlySpan<byte> flags = stackalloc byte[10];
            payload.Write(flags);
            payload.WriteU16(1);
            Biff12RecordWriter.WriteRecord(data, Brt.Xf, payload.Span);
        }

        private ValueTask WriteContentTypesAsync(CancellationToken ct)
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
            return WriteEntryAsync("[Content_Types].xml", sb.ToString(), ct);
        }

        private ValueTask WriteEntryAsync(string entryName, string content, CancellationToken ct)
        {
            return ZipEntryWriter.WriteTextAsync(_zip, entryName, content, _compression, ct);
        }

        private async ValueTask WriteEntryAsync(string entryName, ReadOnlyMemory<byte> content, CancellationToken ct)
        {
            ZipArchiveEntry entry = _zip.CreateEntry(entryName, _compression);
#if NET10_0_OR_GREATER
            var stream = await entry.OpenAsync(ct).ConfigureAwait(false);
#else
            var stream = entry.Open();
#endif
            await using (stream.ConfigureAwait(false))
            {
                await stream.WriteAsync(content, ct).ConfigureAwait(false);
            }
        }
    }
}
