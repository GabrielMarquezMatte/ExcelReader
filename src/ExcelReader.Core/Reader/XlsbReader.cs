using System.Diagnostics.CodeAnalysis;
using System.IO.Compression;

namespace ExcelReader.Core.Reader
{
    // BIFF12 (.xlsb) reader. Uses the same ZIP/OPC container as .xlsx but worksheet parts are
    // binary BIFF12 records. The workbook, styles, and shared-string parts are read once at open
    // time (they're small); worksheets are streamed on demand by the enumerator.
    public sealed partial class XlsbReader : IExcelRowReader, IExcelRowReader<XlsbReader.Enumerator>
    {
        // Shared-string pool: string i = _sharedFlat[_sharedOffsets[i].._sharedOffsets[i+1]].
        private readonly byte[] _sharedFlat = [];
        private readonly int[] _sharedOffsets = [0];
        // Lazily created: dedups repeated shared-string values (categorical columns) into one string
        // instance instead of re-decoding UTF-8 per row. Keyed by the string's stable byte offset into
        // _sharedFlat (see CellDesc.ToCell / Cell.GetString).
        private Dictionary<int, string>? _sharedStringCache;
        private readonly bool[] _styleIsDate = [];
        private readonly ExcelReaderOptions _options;
        private readonly DecompressedByteCounter _decompressedBytes;

        [SuppressMessage("SharpSource", "SS066:Disposable field is not disposed", Justification = "Disposed in Dispose().")]
        private readonly ZipArchive? _zip;
        private readonly Stream? _stream;
        private readonly bool _leaveOpen;
        private readonly (string Name, string Path)[]? _sheets;
        private int _current;

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
        [SuppressMessage("IDisposableAnalyzers.Correctness", "IDISP003:Dispose previous before re-assigning",
            Justification = "Readonly field, first and only assignment in this constructor.")]
        internal XlsbReader(Stream stream, bool leaveOpen, ExcelReaderOptions? options = null)
        {
            _stream = stream;
            _leaveOpen = leaveOpen;
            _options = options ?? ExcelReaderOptions.Default;
            _decompressedBytes = new DecompressedByteCounter(_options.MaxTotalDecompressedBytes);
            _zip = new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen: true);
            try
            {
                var wb = ZipEntryBytes.Read(_zip, "xl/workbook.bin", _decompressedBytes);
                _sheets = XlsbWorkbook.ParseSheets(wb, ZipEntryBytes.Read(_zip, "xl/_rels/workbook.bin.rels", _decompressedBytes));
                if (_sheets.Length == 0)
                {
                    throw new InvalidDataException("The workbook contains no sheets.");
                }
                _styleIsDate = XlsbStyles.ParseStyleDateFlags(ZipEntryBytes.Read(_zip, "xl/styles.bin", _decompressedBytes));
                IsDate1904 = XlsbWorkbook.ParseDate1904(wb);
                (_sharedFlat, _sharedOffsets) = XlsbSharedStrings.Parse(
                    ZipEntryBytes.Read(_zip, "xl/sharedStrings.bin", _decompressedBytes,
                        nameof(ExcelReaderOptions.MaxSharedStringBytes), _options.MaxSharedStringBytes),
                    _options);
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
        }

        [SuppressMessage("IDisposableAnalyzers.Correctness", "IDISP001:Dispose created",
            Justification = "zip ownership transfers to the returned reader on success, disposed in the catch on failure.")]
        internal static async ValueTask<XlsbReader> CreateAsync(Stream stream, bool leaveOpen, ExcelReaderOptions? options = null, CancellationToken ct = default)
        {
            ExcelReaderOptions effectiveOptions = options ?? ExcelReaderOptions.Default;
            DecompressedByteCounter decompressedBytes = new(effectiveOptions.MaxTotalDecompressedBytes);
            ZipArchive? zip = null;
            try
            {
#if NET10_0_OR_GREATER
                zip = await ZipArchive.CreateAsync(stream, ZipArchiveMode.Read, leaveOpen: true, entryNameEncoding: null, ct).ConfigureAwait(false);
#else
                ct.ThrowIfCancellationRequested();
                zip = new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen: true);
#endif
                var wb = await ZipEntryBytes.ReadAsync(zip, "xl/workbook.bin", decompressedBytes, ct).ConfigureAwait(false);
                var zipEntryData = await ZipEntryBytes.ReadAsync(zip, "xl/_rels/workbook.bin.rels", decompressedBytes, ct).ConfigureAwait(false);
                var sheets = XlsbWorkbook.ParseSheets(wb, zipEntryData);
                if (sheets.Length == 0)
                {
                    throw new InvalidDataException("The workbook contains no sheets.");
                }
                var styleIsDate = XlsbStyles.ParseStyleDateFlags(await ZipEntryBytes.ReadAsync(zip, "xl/styles.bin", decompressedBytes, ct).ConfigureAwait(false));
                bool date1904 = XlsbWorkbook.ParseDate1904(wb);
                var (flat, offsets) = XlsbSharedStrings.Parse(
                    await ZipEntryBytes.ReadAsync(zip, "xl/sharedStrings.bin", decompressedBytes, ct,
                        nameof(ExcelReaderOptions.MaxSharedStringBytes), effectiveOptions.MaxSharedStringBytes).ConfigureAwait(false),
                    effectiveOptions);
                return new XlsbReader(stream, leaveOpen, zip, sheets, styleIsDate, date1904, flat, offsets, effectiveOptions, decompressedBytes);
            }
            catch
            {
                if (zip is not null)
                {
#if NET10_0_OR_GREATER
                    await zip.DisposeAsync().ConfigureAwait(false);
#else
                    zip.Dispose();
#endif
                }
                if (!leaveOpen)
                {
                    await stream.DisposeAsync().ConfigureAwait(false);
                }
                throw;
            }
        }

        // --- IExcelReader ---

        public bool IsDate1904 { get; }
        public string SheetName => _sheets![_current].Name;
        public int SheetCount => _sheets!.Length;

        public bool TryMoveToSheet(ReadOnlySpan<char> name)
        {
            if (!WorkbookLookups.TryFindSheetIndex(_sheets!, name, static s => s.Name, out int index))
            {
                return false;
            }
            _current = index;
            return true;
        }

        public void MoveToSheet(int index)
        {
            WorkbookLookups.ValidateSheetIndex(index, _sheets!.Length);
            _current = index;
        }

        // --- Enumeration ---

        internal ReadOnlySpan<byte> SharedSpan => _sharedFlat;

        internal Dictionary<int, string> SharedStringCache => _sharedStringCache ??= [];

        internal bool IsDateStyle(int style)
        {
            return WorkbookLookups.IsDateStyle(_styleIsDate, style);
        }

        internal (int Start, int Length) SharedAt(int index)
        {
            return WorkbookLookups.SharedAt(_sharedOffsets, index);
        }

        [SuppressMessage("Performance", "HLQ006:GetEnumerator should return a value type",
            Justification = "Public nested enumerator is the standard foreach pattern.")]
        public Enumerator GetEnumerator()
        {
            var entry = WorkbookLookups.GetWorksheetEntry(_zip!, _sheets!, _current);
            return new Enumerator(this, WorkbookLookups.OpenEntryStream(entry, _decompressedBytes), entry.Length);
        }

        IExcelRowEnumerator IExcelRowReader<IExcelRowEnumerator>.GetEnumerator()
        {
            return GetEnumerator();
        }

        public async ValueTask<Enumerator> GetAsyncEnumeratorAsync(CancellationToken ct = default)
        {
            var entry = WorkbookLookups.GetWorksheetEntry(_zip!, _sheets!, _current);
#if NET10_0_OR_GREATER
            Stream sheet = new LimitedReadStream(await entry.OpenAsync(ct).ConfigureAwait(false), _decompressedBytes);
#else
            ct.ThrowIfCancellationRequested();
            Stream sheet = WorkbookLookups.OpenEntryStream(entry, _decompressedBytes);
#endif
            return new Enumerator(this, sheet, entry.Length, ct);
        }

        async ValueTask<IExcelRowEnumerator> IExcelRowReader<IExcelRowEnumerator>.GetAsyncEnumeratorAsync(CancellationToken ct)
        {
            return await GetAsyncEnumeratorAsync(ct).ConfigureAwait(false);
        }

        // --- Dispose ---

        [SuppressMessage("IDisposableAnalyzers.Correctness", "IDISP007:Don't dispose injected",
            Justification = "_stream is disposed conditionally based on _leaveOpen.")]
        public void Dispose()
        {
            _zip?.Dispose(); // ZipArchive was opened with leaveOpen:true — does not close _stream
            if (!_leaveOpen)
            {
                _stream?.Dispose();
            }
        }

        [SuppressMessage("IDisposableAnalyzers.Correctness", "IDISP007:Don't dispose injected",
            Justification = "_stream is disposed conditionally based on _leaveOpen.")]
        public async ValueTask DisposeAsync()
        {
            if (_zip is not null)
            {
#if NET10_0_OR_GREATER
                await _zip.DisposeAsync().ConfigureAwait(false);
#else
                _zip.Dispose();
#endif
            }
            if (!_leaveOpen && _stream is not null)
            {
                await _stream.DisposeAsync().ConfigureAwait(false);
            }
        }

    }
}
