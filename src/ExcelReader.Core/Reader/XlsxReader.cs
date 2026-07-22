using System.Buffers;
using System.Diagnostics.CodeAnalysis;
using System.IO.Compression;
using ExcelReader.Core.Writer.Internal;

namespace ExcelReader.Core.Reader
{
    public sealed partial class XlsxReader : IExcelRowReader, IExcelRowReader<XlsxReader.Enumerator>
    {
        private readonly Stream _stream;
        private readonly bool _leaveOpen;
        [SuppressMessage("IDisposableAnalyzers.Correctness", "IDISP008:Don't assign member with injected and created disposables",
            Justification = "Two construction paths: the sync ctor opens and owns the ZipArchive itself; the CreateAsync path receives one already opened by ZipReaderOpen. Either way this reader ends up owning it and disposes it in Dispose/DisposeAsync.")]
        private readonly ZipArchive _zip;
        private readonly ExcelReaderOptions _options;
        private readonly DecompressedByteCounter _decompressedBytes;
        private readonly (string Name, string Path)[] _sheets;
        private readonly bool[] _styleIsDate; // cellXfs index -> true when that style renders as a date/time
        private int _current;

        private byte[] _sharedFlat = [];      // pooled; all decoded shared-string bytes concatenated
        private int[] _sharedOffsets = [0];   // string i = _sharedFlat[_offsets[i].._offsets[i+1]]
        private bool _sharedLoaded;
        // Lazily created: dedups repeated shared-string values (categorical columns) into one string
        // instance instead of re-decoding UTF-8 per row. Keyed by the string's stable byte offset into
        // _sharedFlat (see CellDesc.ToCell / Cell.GetString). Workbook-scoped, so it survives sheet
        // switches (the shared-string table is shared across all sheets in a workbook).
        private Dictionary<int, string>? _sharedStringCache;

        // Sync open: reads the central directory and workbook/styles parts synchronously.
        [SuppressMessage("IDisposableAnalyzers.Correctness", "IDISP003:Dispose previous before re-assigning",
            Justification = "Readonly field, first and only assignment in this constructor.")]
        internal XlsxReader(Stream stream, bool leaveOpen, ExcelReaderOptions? options = null)
        {
            _stream = stream;
            _leaveOpen = leaveOpen;
            _options = options ?? ExcelReaderOptions.Default;
            _decompressedBytes = new DecompressedByteCounter(_options.MaxTotalDecompressedBytes);
            _zip = new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen: true);
            try
            {
                var wbBytes = ZipEntryBytes.Read(_zip, "xl/workbook.xml", _decompressedBytes);
                _sheets = ParseSheets(wbBytes, ZipEntryBytes.Read(_zip, "xl/_rels/workbook.xml.rels", _decompressedBytes));
                if (_sheets.Length == 0)
                {
                    throw new InvalidDataException("The workbook contains no sheets.");
                }
                _styleIsDate = ParseStyleDateFlags(ZipEntryBytes.Read(_zip, "xl/styles.xml", _decompressedBytes));
                IsDate1904 = ParseDate1904(wbBytes);
            }
            catch
            {
                _zip.Dispose();
                if (!leaveOpen)
                {
                    stream.Dispose();
                }
                throw;
            }
        }

        private XlsxReader(Stream stream, bool leaveOpen, ZipArchive zip,
            (string Name, string Path)[] sheets, bool[] styleIsDate, bool date1904,
            ExcelReaderOptions options, DecompressedByteCounter decompressedBytes)
        {
            _stream = stream;
            _leaveOpen = leaveOpen;
            _zip = zip;
            _options = options;
            _decompressedBytes = decompressedBytes;
            _sheets = sheets;
            _styleIsDate = styleIsDate;
            IsDate1904 = date1904;
        }

        // Async open: central directory and parts are read with the .NET 10 async zip APIs.
        internal static ValueTask<XlsxReader> CreateAsync(Stream stream, bool leaveOpen, ExcelReaderOptions? options = null, CancellationToken ct = default)
        {
            ExcelReaderOptions effectiveOptions = options ?? ExcelReaderOptions.Default;
            DecompressedByteCounter decompressedBytes = new(effectiveOptions.MaxTotalDecompressedBytes);
            return ZipReaderOpen.OpenAsync(stream, leaveOpen, async zip =>
            {
                var wb = await ZipEntryBytes.ReadAsync(zip, "xl/workbook.xml", decompressedBytes, ct).ConfigureAwait(false);
                var rels = await ZipEntryBytes.ReadAsync(zip, "xl/_rels/workbook.xml.rels", decompressedBytes, ct).ConfigureAwait(false);
                var sheets = ParseSheets(wb, rels);
                if (sheets.Length == 0)
                {
                    throw new InvalidDataException("The workbook contains no sheets.");
                }
                var styleIsDate = ParseStyleDateFlags(await ZipEntryBytes.ReadAsync(zip, "xl/styles.xml", decompressedBytes, ct).ConfigureAwait(false));
                bool date1904 = ParseDate1904(wb);
                return new XlsxReader(stream, leaveOpen, zip, sheets, styleIsDate, date1904, effectiveOptions, decompressedBytes);
            }, ct);
        }

        public string SheetName => _sheets[_current].Name;
        public int SheetCount => _sheets.Length;
        public bool IsDate1904 { get; }

        // A numeric cell whose style index maps to a date/time format is reported as CellType.Date.
        internal bool IsDateStyle(int style)
        {
            return WorkbookLookups.IsDateStyle(_styleIsDate, style);
        }

        public bool TryMoveToSheet(ReadOnlySpan<char> name)
        {
            if (!WorkbookLookups.TryFindSheetIndex(_sheets, name, static s => s.Name, out int index))
            {
                return false;
            }
            _current = index;
            return true;
        }

        public void MoveToSheet(int index)
        {
            WorkbookLookups.ValidateSheetIndex(index, _sheets.Length);
            _current = index;
        }

        [SuppressMessage("Performance", "HLQ006:GetEnumerator should return a value type",
            Justification = "Enumerator is a class so the same type can also expose MoveNextAsync for the async path.")]
        public Enumerator GetEnumerator()
        {
            EnsureSharedLoaded();
            var entry = WorkbookLookups.GetWorksheetEntry(_zip, _sheets, _current);
            return new Enumerator(this, WorkbookLookups.OpenEntryStream(entry, _decompressedBytes), entry.Length);
        }

        IExcelRowEnumerator IExcelRowReader<IExcelRowEnumerator>.GetEnumerator()
        {
            return GetEnumerator();
        }

        [SuppressMessage("Performance", "HLQ006:GetAsyncEnumerator should return a value type",
            Justification = "Enumerator is a class so the same type can also expose MoveNextAsync for the async path.")]
        public Enumerator GetAsyncEnumerator()
        {
            return GetEnumerator();
        }

        IExcelRowEnumerator IExcelRowReader<IExcelRowEnumerator>.GetAsyncEnumerator()
        {
            return GetAsyncEnumerator();
        }

        /// <summary>
        /// Streaming async enumerator over the current sheet. Use with a manual loop — <c>Current</c>
        /// is a ref struct (<c>Row</c>), so <c>await foreach</c> cannot bind it:
        /// <code>
        /// await using var e = await reader.GetAsyncEnumeratorAsync(ct);
        /// while (await e.MoveNextAsync()) { var row = e.Current; /* ... */ }
        /// </code>
        /// </summary>
        [SuppressMessage("IDisposableAnalyzers.Correctness", "IDISP001:Dispose created",
            Justification = "sheet is handed to the returned Enumerator, which owns and disposes it.")]
        public async ValueTask<Enumerator> GetAsyncEnumeratorAsync(CancellationToken ct = default)
        {
            await EnsureSharedLoadedAsync(ct).ConfigureAwait(false);
            var entry = WorkbookLookups.GetWorksheetEntry(_zip, _sheets, _current);
#if NET10_0_OR_GREATER
            var sheet = new LimitedReadStream(await entry.OpenAsync(ct).ConfigureAwait(false), _decompressedBytes);
#else
            ct.ThrowIfCancellationRequested();
            var sheet = WorkbookLookups.OpenEntryStream(entry, _decompressedBytes);
#endif
            return new Enumerator(this, sheet, entry.Length, ct);
        }

        async ValueTask<IExcelRowEnumerator> IExcelRowReader<IExcelRowEnumerator>.GetAsyncEnumeratorAsync(CancellationToken ct)
        {
            return await GetAsyncEnumeratorAsync(ct).ConfigureAwait(false);
        }

        internal ReadOnlySpan<byte> SharedSpan => _sharedFlat;

        internal Dictionary<int, string> SharedStringCache => _sharedStringCache ??= new Dictionary<int, string>();

        internal (int Start, int Length) SharedAt(int index)
        {
            return WorkbookLookups.SharedAt(_sharedOffsets, index);
        }

        public void Dispose()
        {
            if (_sharedFlat.Length > 0)
            {
                ArrayPool<byte>.Shared.Return(_sharedFlat);
                _sharedFlat = [];
            }
            _zip.Dispose();
            if (!_leaveOpen)
            {
                _stream.Dispose();
            }
        }

        public async ValueTask DisposeAsync()
        {
            if (_sharedFlat.Length > 0)
            {
                ArrayPool<byte>.Shared.Return(_sharedFlat);
                _sharedFlat = [];
            }
            await ZipArchiveDisposal.DisposeAsync(_zip).ConfigureAwait(false);
            if (!_leaveOpen)
            {
                await _stream.DisposeAsync().ConfigureAwait(false);
            }
        }
    }
}
