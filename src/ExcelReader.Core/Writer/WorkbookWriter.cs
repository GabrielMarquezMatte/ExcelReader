using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.IO.Compression;
using System.Text;
using ExcelReader.Core.Writer.Internal;

namespace ExcelReader.Core.Writer
{
    public sealed class WorkbookWriter : IAsyncDisposable
    {
        private readonly ZipArchive _zip;
        private readonly Stream _stream;
        private readonly bool _leaveOpen;
        private readonly CompressionLevel _compression;
        private readonly List<(string Name, int SheetId)> _sheets = [];
        private WriterState _state = WriterState.Created;
        private bool _sheetActive;
        private SheetWriter? _activeSheet;
        private bool _disposed;

        private WorkbookWriter(ZipArchive zip, Stream stream, bool leaveOpen, CompressionLevel compression)
        {
            _zip = zip;
            _stream = stream;
            _leaveOpen = leaveOpen;
            _compression = compression;
        }

        [SuppressMessage("IDisposableAnalyzers.Correctness", "IDISP001:Dispose created",
            Justification = "Factory method transfers ZipArchive ownership to WorkbookWriter; caller disposes via DisposeAsync.")]
        [SuppressMessage("Reliability", "CA2000:Dispose objects before losing scope",
            Justification = "WorkbookWriter takes ownership of ZipArchive and disposes it in DisposeAsync/EndAsync.")]
        public static ValueTask<WorkbookWriter> CreateAsync(
            Stream stream, bool leaveOpen = false,
            CompressionLevel compression = CompressionLevel.Fastest, CancellationToken ct = default)
        {
            ArgumentNullException.ThrowIfNull(stream);
            ct.ThrowIfCancellationRequested();
            ZipArchive zip = new(stream, ZipArchiveMode.Create, leaveOpen: true);
            return ValueTask.FromResult(new WorkbookWriter(zip, stream, leaveOpen, compression));
        }

        public ValueTask StartAsync(CancellationToken ct = default)
        {
            ObjectDisposedException.ThrowIf(_state == WriterState.Ended, this);
            if (_state != WriterState.Created)
            {
                throw new InvalidOperationException("WorkbookWriter has already been started.");
            }
            ct.ThrowIfCancellationRequested();
            _state = WriterState.Started;
            return WriteRootRelsAsync(ct);
        }

        public SheetWriter AddSheet(string name)
        {
            ArgumentNullException.ThrowIfNull(name);
            ObjectDisposedException.ThrowIf(_state == WriterState.Ended, this);
            if (_state != WriterState.Started)
            {
                throw new InvalidOperationException("WorkbookWriter must be started before adding sheets.");
            }
            if (_sheetActive)
            {
                throw new InvalidOperationException("The previous SheetWriter must be ended before adding a new sheet.");
            }
            _sheetActive = true;
            int sheetId = _sheets.Count + 1;
            _activeSheet = new SheetWriter(this, _zip, name, sheetId, _compression);
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

        public async ValueTask EndAsync(CancellationToken ct = default)
        {
            ObjectDisposedException.ThrowIf(_state == WriterState.Ended, this);
            if (_state != WriterState.Started)
            {
                throw new InvalidOperationException("WorkbookWriter must be started before ending.");
            }
            ct.ThrowIfCancellationRequested();
            _state = WriterState.Ended;
            if (_activeSheet is not null)
            {
                await _activeSheet.DisposeAsync().ConfigureAwait(false);
            }
            await WriteStylesAsync(ct).ConfigureAwait(false);
            await WriteWorkbookAsync(ct).ConfigureAwait(false);
            await WriteWorkbookRelsAsync(ct).ConfigureAwait(false);
            await WriteContentTypesAsync(ct).ConfigureAwait(false);
            await _zip.DisposeAsync().ConfigureAwait(false);
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
                await EndAsync().ConfigureAwait(false);
            }
            else if (_state == WriterState.Created)
            {
                await _zip.DisposeAsync().ConfigureAwait(false);
            }
            if (!_leaveOpen)
            {
                await _stream.DisposeAsync().ConfigureAwait(false);
            }
        }

        [SuppressMessage("IDisposableAnalyzers.Correctness", "IDISP001:Dispose created",
            Justification = "Stream is disposed indirectly by StreamWriter with leaveOpen: false.")]
        private async ValueTask WriteRootRelsAsync(CancellationToken ct)
        {
            ZipArchiveEntry entry = _zip.CreateEntry("_rels/.rels", _compression);
            Stream stream = await entry.OpenAsync(ct).ConfigureAwait(false);
            StreamWriter xml = new(stream, new UTF8Encoding(encoderShouldEmitUTF8Identifier: false), leaveOpen: false);
            await using (xml.ConfigureAwait(false))
            {
                await xml.WriteAsync(
                    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                    $"<Relationships xmlns=\"{XlsxConstants.PackageRelationshipsNs}\">" +
                    $"<Relationship Id=\"rId1\" Type=\"{XlsxConstants.WorkbookRelType}\" Target=\"xl/workbook.xml\"/>" +
                    "</Relationships>").ConfigureAwait(false);
                await xml.FlushAsync(ct).ConfigureAwait(false);
            }
        }

        [SuppressMessage("IDisposableAnalyzers.Correctness", "IDISP001:Dispose created",
            Justification = "Stream is disposed indirectly by StreamWriter with leaveOpen: false.")]
        private async ValueTask WriteStylesAsync(CancellationToken ct)
        {
            ZipArchiveEntry entry = _zip.CreateEntry("xl/styles.xml", _compression);
            Stream stream = await entry.OpenAsync(ct).ConfigureAwait(false);
            StreamWriter xml = new(stream, new UTF8Encoding(encoderShouldEmitUTF8Identifier: false), leaveOpen: false);
            await using (xml.ConfigureAwait(false))
            {
                await xml.WriteAsync(
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
                    "</styleSheet>").ConfigureAwait(false);
                await xml.FlushAsync(ct).ConfigureAwait(false);
            }
        }

        [SuppressMessage("IDisposableAnalyzers.Correctness", "IDISP001:Dispose created",
            Justification = "Stream is disposed indirectly by StreamWriter with leaveOpen: false.")]
        private async ValueTask WriteWorkbookAsync(CancellationToken ct)
        {
            ZipArchiveEntry entry = _zip.CreateEntry("xl/workbook.xml", _compression);
            Stream stream = await entry.OpenAsync(ct).ConfigureAwait(false);
            StreamWriter xml = new(stream, new UTF8Encoding(encoderShouldEmitUTF8Identifier: false), leaveOpen: false);
            await using (xml.ConfigureAwait(false))
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
                await xml.WriteAsync(sb.ToString()).ConfigureAwait(false);
                await xml.FlushAsync(ct).ConfigureAwait(false);
            }
        }

        [SuppressMessage("IDisposableAnalyzers.Correctness", "IDISP001:Dispose created",
            Justification = "Stream is disposed indirectly by StreamWriter with leaveOpen: false.")]
        private async ValueTask WriteWorkbookRelsAsync(CancellationToken ct)
        {
            ZipArchiveEntry entry = _zip.CreateEntry("xl/_rels/workbook.xml.rels", _compression);
            Stream stream = await entry.OpenAsync(ct).ConfigureAwait(false);
            StreamWriter xml = new(stream, new UTF8Encoding(encoderShouldEmitUTF8Identifier: false), leaveOpen: false);
            await using (xml.ConfigureAwait(false))
            {
                StringBuilder sb = new();
                sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
                sb.Append($"<Relationships xmlns=\"{XlsxConstants.PackageRelationshipsNs}\">");
                sb.Append($"<Relationship Id=\"rId1\" Type=\"{XlsxConstants.StylesRelType}\" Target=\"styles.xml\"/>");
                foreach ((_, int sheetId) in _sheets)
                {
                    sb.Append(CultureInfo.InvariantCulture, $"<Relationship Id=\"rId{sheetId + 1}\" Type=\"{XlsxConstants.WorksheetRelType}\" Target=\"worksheets/sheet{sheetId}.xml\"/>");
                }
                sb.Append("</Relationships>");
                await xml.WriteAsync(sb.ToString()).ConfigureAwait(false);
                await xml.FlushAsync(ct).ConfigureAwait(false);
            }
        }

        [SuppressMessage("IDisposableAnalyzers.Correctness", "IDISP001:Dispose created",
            Justification = "Stream is disposed indirectly by StreamWriter with leaveOpen: false.")]
        private async ValueTask WriteContentTypesAsync(CancellationToken ct)
        {
            ZipArchiveEntry entry = _zip.CreateEntry("[Content_Types].xml", _compression);
            Stream stream = await entry.OpenAsync(ct).ConfigureAwait(false);
            StreamWriter xml = new(stream, new UTF8Encoding(encoderShouldEmitUTF8Identifier: false), leaveOpen: false);
            await using (xml.ConfigureAwait(false))
            {
                StringBuilder sb = new();
                sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
                sb.Append($"<Types xmlns=\"{XlsxConstants.ContentTypesNs}\">");
                sb.Append($"<Default Extension=\"rels\" ContentType=\"{XlsxConstants.RelationshipsContentType}\"/>");
                sb.Append("<Default Extension=\"xml\" ContentType=\"application/xml\"/>");
                sb.Append($"<Override PartName=\"/xl/workbook.xml\" ContentType=\"{XlsxConstants.WorkbookContentType}\"/>");
                sb.Append($"<Override PartName=\"/xl/styles.xml\" ContentType=\"{XlsxConstants.StylesContentType}\"/>");
                foreach ((_, int sheetId) in _sheets)
                {
                    sb.Append(CultureInfo.InvariantCulture, $"<Override PartName=\"/xl/worksheets/sheet{sheetId}.xml\" ContentType=\"{XlsxConstants.WorksheetContentType}\"/>");
                }
                sb.Append("</Types>");
                await xml.WriteAsync(sb.ToString()).ConfigureAwait(false);
                await xml.FlushAsync(ct).ConfigureAwait(false);
            }
        }

        private static string EscapeAttribute(string value)
        {
            StringBuilder sb = new(value);
            return sb.Replace("&", "&amp;")
                     .Replace("<", "&lt;")
                     .Replace(">", "&gt;")
                     .Replace("\"", "&quot;")
                     .ToString();
        }
    }
}
