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
        private static readonly UTF8Encoding Utf8NoBom = new(encoderShouldEmitUTF8Identifier: false);
        private const int MaxSheetNameLength = 31;

        private readonly ZipArchive _zip;
        private readonly Stream _stream;
        private readonly bool _leaveOpen;
        private readonly bool _date1904;
        private readonly CompressionLevel _compression;
        private readonly List<XlsbSheetWriter> _sheets = [];
        private WriterState _state = WriterState.Created;
        private XlsbSheetWriter? _activeSheet;
        private bool _disposed;

        private XlsbWorkbookWriter(ZipArchive zip, Stream stream, bool leaveOpen, bool date1904, CompressionLevel compression)
        {
            _zip = zip;
            _stream = stream;
            _leaveOpen = leaveOpen;
            _date1904 = date1904;
            _compression = compression;
        }

        [SuppressMessage("IDisposableAnalyzers.Correctness", "IDISP001:Dispose created",
            Justification = "Factory method transfers ZipArchive ownership to XlsbWorkbookWriter; caller disposes via DisposeAsync.")]
        [SuppressMessage("Reliability", "CA2000:Dispose objects before losing scope",
            Justification = "XlsbWorkbookWriter takes ownership of ZipArchive and disposes it in DisposeAsync/EndAsync.")]
        public static ValueTask<XlsbWorkbookWriter> CreateAsync(
            Stream stream,
            bool leaveOpen = false,
            bool date1904 = false,
            CompressionLevel compression = CompressionLevel.Fastest,
            CancellationToken ct = default)
        {
            ArgumentNullException.ThrowIfNull(stream);
            ct.ThrowIfCancellationRequested();
            ZipArchive zip = new(stream, ZipArchiveMode.Create, leaveOpen: true);
            return ValueTask.FromResult(new XlsbWorkbookWriter(zip, stream, leaveOpen, date1904, compression));
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
            _activeSheet = new XlsbSheetWriter(this, name, sheetId, _date1904);
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
            await WriteEntryAsync("xl/sharedStrings.bin", ReadOnlyMemory<byte>.Empty, ct).ConfigureAwait(false);
            foreach (XlsbSheetWriter sheet in _sheets)
            {
                await WriteEntryAsync($"xl/worksheets/sheet{sheet.SheetId}.bin", sheet.Memory, ct).ConfigureAwait(false);
            }
            await WriteContentTypesAsync(ct).ConfigureAwait(false);
#if NET10_0_OR_GREATER
            await _zip.DisposeAsync().ConfigureAwait(false);
#else
            _zip.Dispose();
#endif
            foreach (ref readonly var sheet in CollectionsMarshal.AsSpan(_sheets))
            {
                sheet.ReleaseBuffer();
            }
        }

        public ValueTask FlushAsync(CancellationToken ct = default)
        {
            ct.ThrowIfCancellationRequested();
            return new ValueTask(_stream.FlushAsync(ct));
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
                $"<Relationship Id=\"rId1\" Type=\"{XlsxConstants.WorkbookRelType}\" Target=\"xl/workbook.bin\"/>" +
                "</Relationships>";
            return WriteEntryAsync("_rels/.rels", xml, ct);
        }

        private async ValueTask WriteWorkbookAsync(CancellationToken ct)
        {
            using BiffBuffer payload = new(256);
            using BiffBuffer data = new(1024);
            payload.WriteU32(_date1904 ? 1u : 0u);
            payload.WriteU32(0);
            Biff12RecordWriter.WriteWideString(payload, string.Empty);
            Biff12RecordWriter.WriteRecord(data, Brt.WbProp, payload.Span);

            foreach (ref readonly var sheet in CollectionsMarshal.AsSpan(_sheets))
            {
                payload.Reset();
                payload.WriteU32(0);
                payload.WriteU32((uint)sheet.SheetId);
                Biff12RecordWriter.WriteWideString(payload, $"rId{sheet.SheetId}");
                Biff12RecordWriter.WriteWideString(payload, sheet.Name);
                Biff12RecordWriter.WriteRecord(data, Brt.BundleSh, payload.Span);
            }
            await WriteEntryAsync("xl/workbook.bin", data.Memory, ct).ConfigureAwait(false);
        }

        private ValueTask WriteWorkbookRelsAsync(CancellationToken ct)
        {
            StringBuilder sb = new();
            sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
            sb.Append(CultureInfo.InvariantCulture, $"<Relationships xmlns=\"{XlsxConstants.PackageRelationshipsNs}\">");
            sb.Append(CultureInfo.InvariantCulture, $"<Relationship Id=\"rIdStyles\" Type=\"{XlsxConstants.StylesRelType}\" Target=\"styles.bin\"/>");
            sb.Append(CultureInfo.InvariantCulture, $"<Relationship Id=\"rIdShared\" Type=\"{XlsxConstants.SharedStringsRelType}\" Target=\"sharedStrings.bin\"/>");
            foreach (ref readonly var sheet in CollectionsMarshal.AsSpan(_sheets))
            {
                sb.Append(CultureInfo.InvariantCulture, $"<Relationship Id=\"rId{sheet.SheetId}\" Type=\"{XlsxConstants.WorksheetRelType}\" Target=\"worksheets/sheet{sheet.SheetId}.bin\"/>");
            }
            sb.Append("</Relationships>");
            return WriteEntryAsync("xl/_rels/workbook.bin.rels", sb.ToString(), ct);
        }

        private async ValueTask WriteStylesAsync(CancellationToken ct)
        {
            using BiffBuffer payload = new(32);
            using BiffBuffer data = new(128);
            Biff12RecordWriter.WriteRecord(data, Brt.BeginCellXFs);
            WriteXf(data, payload, 0);
            WriteXf(data, payload, 14);
            Biff12RecordWriter.WriteRecord(data, Brt.EndCellXFs);
            await WriteEntryAsync("xl/styles.bin", data.Memory, ct).ConfigureAwait(false);
        }

        private static void WriteXf(BiffBuffer data, BiffBuffer payload, int numFmtId)
        {
            payload.Reset();
            payload.WriteU16(0);
            payload.WriteU16(numFmtId);
            payload.Write(new byte[12]);
            Biff12RecordWriter.WriteRecord(data, Brt.Xf, payload.Span);
        }

        private ValueTask WriteContentTypesAsync(CancellationToken ct)
        {
            StringBuilder sb = new();
            sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
            sb.Append(CultureInfo.InvariantCulture, $"<Types xmlns=\"{XlsxConstants.ContentTypesNs}\">");
            sb.Append(CultureInfo.InvariantCulture, $"<Default Extension=\"rels\" ContentType=\"{XlsxConstants.RelationshipsContentType}\"/>");
            sb.Append("<Default Extension=\"bin\" ContentType=\"application/vnd.ms-excel.sheet.binary.macroEnabled.main\"/>");
            sb.Append("<Override PartName=\"/xl/workbook.bin\" ContentType=\"application/vnd.ms-excel.sheet.binary.macroEnabled.main\"/>");
            sb.Append("<Override PartName=\"/xl/styles.bin\" ContentType=\"application/vnd.ms-excel.styles\"/>");
            sb.Append("<Override PartName=\"/xl/sharedStrings.bin\" ContentType=\"application/vnd.ms-excel.sharedStrings\"/>");
            foreach (ref readonly var sheet in CollectionsMarshal.AsSpan(_sheets))
            {
                sb.Append(CultureInfo.InvariantCulture, $"<Override PartName=\"/xl/worksheets/sheet{sheet.SheetId}.bin\" ContentType=\"application/vnd.ms-excel.worksheet\"/>");
            }
            sb.Append("</Types>");
            return WriteEntryAsync("[Content_Types].xml", sb.ToString(), ct);
        }

        private async ValueTask WriteEntryAsync(string entryName, string content, CancellationToken ct)
        {
            ZipArchiveEntry entry = _zip.CreateEntry(entryName, _compression);
#if NET10_0_OR_GREATER
            var stream = await entry.OpenAsync(ct).ConfigureAwait(false);
#else
            Stream stream = entry.Open();
#endif
            StreamWriter writer = new(stream, Utf8NoBom, leaveOpen: false);
            await using (stream.ConfigureAwait(false))
            await using (writer.ConfigureAwait(false))
            {
                await writer.WriteAsync(content).ConfigureAwait(false);
                await writer.FlushAsync(ct).ConfigureAwait(false);
            }
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