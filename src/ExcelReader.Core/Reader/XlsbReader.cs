using System.Buffers;
using System.Diagnostics.CodeAnalysis;
using System.IO.Compression;
using ExcelReader.Core.Internal;

namespace ExcelReader.Core.Reader
{
    /// <summary>Reads rows from a binary Excel (.xlsb / BIFF12) workbook, streaming each sheet's cells without loading the whole file into memory.</summary>
    /// <remarks>
    /// Uses the same ZIP/OPC container as .xlsx, but worksheet parts are binary BIFF12 records. The workbook,
    /// styles, and shared-string parts are read once at open time (they're small); worksheets are streamed on
    /// demand by the enumerator.
    /// </remarks>
    public sealed partial class XlsbReader : IExcelRowReader, IExcelRowReader<XlsbReader.Enumerator>
    {
        // Shared-string pool: string i = _sharedFlat[_sharedOffsets[i].._sharedOffsets[i+1]].
        private readonly byte[] _sharedFlat = [];
        private readonly int[] _sharedOffsets = [0];
        private readonly bool _pooledSharedFlat;
        // Lazily created: dedups repeated shared-string values (categorical columns) into one string
        // instance instead of re-decoding UTF-8 per row. Indexed by shared-string index (see
        // WorkbookLookups.CreateSharedStringCache, CellDesc.ToCell, Cell.GetString).
        private string?[]? _sharedStringCache;
        private readonly bool[] _styleIsDate = [];
        private readonly ExcelReaderOptions _options;
        private readonly DecompressedByteCounter _decompressedBytes;

        [SuppressMessage("IDisposableAnalyzers.Correctness", "IDISP008:Don't assign member with injected and created disposables",
            Justification = "Two construction paths: the sync ctor opens and owns the ZipArchive itself; the CreateAsync path receives one already opened by ZipReaderOpen. Either way this reader ends up owning it and disposes it in Dispose/DisposeAsync.")]
        private readonly ZipArchive? _zip;
        // Non-null instead of _zip/_stream for the in-memory ZIP path — exactly one of _zip or _memZip
        // is non-null for any reader instance other than the test-only ctor.
        private readonly ZipMemoryIndex? _memZip;
        private readonly Stream? _stream;
        private readonly bool _leaveOpen;
        private readonly (string Name, string Path)[]? _sheets;
        private int _current;
        private int _disposed;

        // Test-only: accepts pre-parsed components (no ZIP, no stream navigation).
        internal XlsbReader(byte[] sharedFlat, int[] sharedOffsets, bool[] styleIsDate, bool date1904)
        {
            _options = ExcelReaderOptions.Default;
            _decompressedBytes = new DecompressedByteCounter(_options.MaxTotalDecompressedBytes);
            _sharedFlat = sharedFlat;
            _sharedOffsets = sharedOffsets;
            _styleIsDate = styleIsDate;
            IsDate1904 = date1904;
        }

        // Sync open: reads the three small workbook parts, keeps _zip open for worksheet streaming.
        internal XlsbReader(Stream stream, bool leaveOpen, ExcelReaderOptions? options = null)
            : this(stream, leaveOpen, new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen: true), options)
        {
        }

        // Sync open over an already-opened ZipArchive — lets a caller that already opened the archive
        // for format detection (Excel.Open's DetectSeekable) hand it straight to the reader instead of
        // re-parsing the central directory a second time.
        internal XlsbReader(Stream stream, bool leaveOpen, ZipArchive zip, ExcelReaderOptions? options = null)
        {
            _stream = stream;
            _leaveOpen = leaveOpen;
            _options = options ?? ExcelReaderOptions.Default;
            _decompressedBytes = new DecompressedByteCounter(_options.MaxTotalDecompressedBytes);
            _zip = zip;
            try
            {
                LimitChecks.ThrowIfTooManyEntries(_zip.Entries.Count, _options);
                var wb = ZipEntryBytes.Read(_zip, "xl/workbook.bin", _decompressedBytes);
                _sheets = XlsbWorkbook.ParseSheets(wb, ZipEntryBytes.Read(_zip, "xl/_rels/workbook.bin.rels", _decompressedBytes));
                if (_sheets.Length == 0)
                {
                    throw new InvalidDataException("The workbook contains no sheets.");
                }
                _styleIsDate = XlsbStyles.ParseStyleDateFlags(ZipEntryBytes.Read(_zip, "xl/styles.bin", _decompressedBytes));
                IsDate1904 = XlsbWorkbook.ParseDate1904(wb);
                (_sharedFlat, _sharedOffsets) = LoadSharedStrings(_zip);
                _pooledSharedFlat = _sharedFlat.Length != 0;
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

        // Private constructor used by CreateAsync — all parts already parsed.
        private XlsbReader(Stream stream, bool leaveOpen, ZipArchive zip,
            (string Name, string Path)[] sheets, bool[] styleIsDate, bool date1904,
            byte[] sharedFlat, int[] sharedOffsets, ExcelReaderOptions options, DecompressedByteCounter decompressedBytes)
        {
            _stream = stream;
            _leaveOpen = leaveOpen;
            _zip = zip;
            _options = options;
            _decompressedBytes = decompressedBytes;
            _sheets = sheets;
            _styleIsDate = styleIsDate;
            IsDate1904 = date1904;
            _sharedFlat = sharedFlat;
            _sharedOffsets = sharedOffsets;
            _pooledSharedFlat = sharedFlat.Length != 0;
        }

        // In-memory ZIP path: no stream, no ZipArchive. sharedFlat here
        // comes from XlsbSharedStrings.Parse (a plain array, not ArrayPool-rented), unlike the streamed
        // ctor above — _pooledSharedFlat is always false.
        private XlsbReader(ZipMemoryIndex memZip,
            (string Name, string Path)[] sheets, bool[] styleIsDate, bool date1904,
            byte[] sharedFlat, int[] sharedOffsets, ExcelReaderOptions options, DecompressedByteCounter decompressedBytes)
        {
            _leaveOpen = true;
            _memZip = memZip;
            _options = options;
            _decompressedBytes = decompressedBytes;
            _sheets = sheets;
            _styleIsDate = styleIsDate;
            IsDate1904 = date1904;
            _sharedFlat = sharedFlat;
            _sharedOffsets = sharedOffsets;
            _pooledSharedFlat = false;
        }

        internal static ValueTask<XlsbReader> CreateAsync(Stream stream, bool leaveOpen, ExcelReaderOptions? options = null, CancellationToken ct = default)
        {
            ExcelReaderOptions effectiveOptions = options ?? ExcelReaderOptions.Default;
            DecompressedByteCounter decompressedBytes = new(effectiveOptions.MaxTotalDecompressedBytes);
            return ZipReaderOpen.OpenAsync(stream, leaveOpen, effectiveOptions,
                zip => ParseAsync(stream, leaveOpen, zip, effectiveOptions, decompressedBytes, ct), ct);
        }

        // Async open over an already-opened ZipArchive — the async twin of the ZipArchive-taking sync
        // ctor above, for callers (Excel.OpenAsync's DetectSeekableAsync) that already opened the
        // archive for format detection.
        internal static ValueTask<XlsbReader> CreateFromOpenZipAsync(
            Stream stream, bool leaveOpen, ZipArchive zip, ExcelReaderOptions? options, CancellationToken ct)
        {
            ExcelReaderOptions effectiveOptions = options ?? ExcelReaderOptions.Default;
            DecompressedByteCounter decompressedBytes = new(effectiveOptions.MaxTotalDecompressedBytes);
            return ZipReaderOpen.FromOpenZipAsync(stream, leaveOpen, zip, effectiveOptions,
                z => ParseAsync(stream, leaveOpen, z, effectiveOptions, decompressedBytes, ct));
        }

        private static async ValueTask<XlsbReader> ParseAsync(
            Stream stream, bool leaveOpen, ZipArchive zip, ExcelReaderOptions effectiveOptions,
            DecompressedByteCounter decompressedBytes, CancellationToken ct)
        {
            var wb = await ZipEntryBytes.ReadAsync(zip, "xl/workbook.bin", decompressedBytes, ct).ConfigureAwait(false);
            var zipEntryData = await ZipEntryBytes.ReadAsync(zip, "xl/_rels/workbook.bin.rels", decompressedBytes, ct).ConfigureAwait(false);
            var sheets = XlsbWorkbook.ParseSheets(wb, zipEntryData);
            if (sheets.Length == 0)
            {
                throw new InvalidDataException("The workbook contains no sheets.");
            }
            var styleIsDate = XlsbStyles.ParseStyleDateFlags(await ZipEntryBytes.ReadAsync(zip, "xl/styles.bin", decompressedBytes, ct).ConfigureAwait(false));
            bool date1904 = XlsbWorkbook.ParseDate1904(wb);
            var (flat, offsets) = await LoadSharedStringsAsync(zip, decompressedBytes, effectiveOptions, ct).ConfigureAwait(false);
            return new XlsbReader(stream, leaveOpen, zip, sheets, styleIsDate, date1904, flat, offsets, effectiveOptions, decompressedBytes);
        }

        private (byte[] Flat, int[] Offsets) LoadSharedStrings(ZipArchive zip)
        {
            ZipArchiveEntry? entry = zip.GetEntry("xl/sharedStrings.bin");
            if (entry is null)
            {
                return ([], [0]);
            }
            WorkbookLookups.ThrowIfSharedEntryTooLarge(entry.Length, _decompressedBytes, _options);
            using LimitedReadStream stream = WorkbookLookups.OpenEntryStream(entry, _decompressedBytes, _options,
                nameof(ExcelReaderOptions.MaxSharedStringBytes), _options.MaxSharedStringBytes);
            return XlsbSharedStrings.ParseStreaming(stream, entry.Length, _options);
        }

        private static async ValueTask<(byte[] Flat, int[] Offsets)> LoadSharedStringsAsync(
            ZipArchive zip, DecompressedByteCounter decompressedBytes, ExcelReaderOptions options, CancellationToken ct)
        {
            ZipArchiveEntry? entry = zip.GetEntry("xl/sharedStrings.bin");
            if (entry is null)
            {
                return ([], [0]);
            }
            WorkbookLookups.ThrowIfSharedEntryTooLarge(entry.Length, decompressedBytes, options);
            LimitedReadStream stream = await WorkbookLookups.OpenEntryStreamAsync(entry, decompressedBytes, options, ct,
                nameof(ExcelReaderOptions.MaxSharedStringBytes), options.MaxSharedStringBytes).ConfigureAwait(false);
            await using (stream.ConfigureAwait(false))
            {
                return await XlsbSharedStrings.ParseStreamingAsync(stream, entry.Length, options, ct).ConfigureAwait(false);
            }
        }

        // --- IExcelReader ---

        /// <inheritdoc/>
        public bool IsDate1904 { get; }
        /// <inheritdoc/>
        public string SheetName => _sheets![_current].Name;
        /// <inheritdoc/>
        public int SheetCount => _sheets!.Length;

        /// <inheritdoc/>
        public bool TryMoveToSheet(ReadOnlySpan<char> name)
        {
            if (!WorkbookLookups.TryFindSheetIndex(_sheets!, name, static s => s.Name, out int index))
            {
                return false;
            }
            _current = index;
            return true;
        }

        /// <inheritdoc/>
        public void MoveToSheet(int index)
        {
            WorkbookLookups.ValidateSheetIndex(index, _sheets!.Length);
            _current = index;
        }

        // --- Enumeration ---

        internal ReadOnlySpan<byte> SharedSpan => _sharedFlat;

        internal string?[] SharedStringCache => _sharedStringCache ??= WorkbookLookups.CreateSharedStringCache(_sharedOffsets);

        internal bool IsDateStyle(int style)
        {
            return WorkbookLookups.IsDateStyle(_styleIsDate, style);
        }

        /// <inheritdoc/>
        [SuppressMessage("Performance", "HLQ006:GetEnumerator should return a value type",
            Justification = "Public nested enumerator is the standard foreach pattern.")]
        public Enumerator GetEnumerator()
        {
            if (_memZip is not null)
            {
                return GetEnumeratorFromMemory();
            }
            var entry = WorkbookLookups.GetWorksheetEntry(_zip!, _sheets!, _current);
            return new Enumerator(this, WorkbookLookups.OpenEntryStream(entry, _decompressedBytes, _options), entry.Length);
        }

        IExcelRowEnumerator IExcelRowReader<IExcelRowEnumerator>.GetEnumerator()
        {
            return GetEnumerator();
        }

        /// <inheritdoc/>
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
        [SuppressMessage("Reliability", "CA2000:Dispose objects before losing scope",
            Justification = "The enumerator is handed to the caller via the returned ValueTask, which owns and disposes it.")]
        public ValueTask<Enumerator> GetAsyncEnumeratorAsync(CancellationToken ct = default)
        {
            if (_memZip is not null)
            {
                return new ValueTask<Enumerator>(GetEnumeratorFromMemory());
            }
            return GetAsyncEnumeratorFromStreamAsync(ct);
        }

        [SuppressMessage("IDisposableAnalyzers.Correctness", "IDISP001:Dispose created",
            Justification = "sheet is handed to the returned Enumerator, which owns and disposes it.")]
        private async ValueTask<Enumerator> GetAsyncEnumeratorFromStreamAsync(CancellationToken ct)
        {
            var entry = WorkbookLookups.GetWorksheetEntry(_zip!, _sheets!, _current);
            Stream sheet = await WorkbookLookups
                .OpenEntryStreamAsync(entry, _decompressedBytes, _options, ct).ConfigureAwait(false);
            return new Enumerator(this, sheet, entry.Length, ct);
        }

        async ValueTask<IExcelRowEnumerator> IExcelRowReader<IExcelRowEnumerator>.GetAsyncEnumeratorAsync(CancellationToken ct)
        {
            return await GetAsyncEnumeratorAsync(ct).ConfigureAwait(false);
        }

        // --- Dispose ---

        /// <inheritdoc/>
        [SuppressMessage("IDisposableAnalyzers.Correctness", "IDISP007:Don't dispose injected",
            Justification = "_memZip's lifetime was transferred to this reader at construction; this is owned disposal, not disposing a borrowed dependency.")]
        public void Dispose()
        {
            if (Interlocked.Exchange(ref _disposed, 1) != 0)
            {
                return;
            }
            _memZip?.Dispose();
            _zip?.Dispose(); // ZipArchive was opened with leaveOpen:true — does not close _stream
            if (!_leaveOpen)
            {
                _stream?.Dispose();
            }
            if (_pooledSharedFlat)
            {
                ArrayPool<byte>.Shared.Return(_sharedFlat);
            }
        }

        /// <inheritdoc/>
        [SuppressMessage("IDisposableAnalyzers.Correctness", "IDISP007:Don't dispose injected",
            Justification = "_stream is disposed conditionally based on _leaveOpen.")]
        public async ValueTask DisposeAsync()
        {
            if (Interlocked.Exchange(ref _disposed, 1) != 0)
            {
                return;
            }
            _memZip?.Dispose();
            if (_zip is not null)
            {
                await ZipArchiveDisposal.DisposeAsync(_zip).ConfigureAwait(false);
            }
            if (!_leaveOpen && _stream is not null)
            {
                await _stream.DisposeAsync().ConfigureAwait(false);
            }
            if (_pooledSharedFlat)
            {
                ArrayPool<byte>.Shared.Return(_sharedFlat);
            }
        }

    }
}
