using System.Buffers;
using System.IO.Compression;
using ExcelReader.Core.Enums;

namespace ExcelReader.Core.Reader
{
    /// <summary>Reads rows from an Office Open XML (.xlsx) workbook, streaming each sheet's cells without loading the whole file into memory.</summary>
    public sealed partial class XlsxReader : IExcelRowReader, IExcelRowReader<XlsxReader.Enumerator>
    {
        private readonly Stream? _stream;
        private readonly bool _leaveOpen;
        private readonly ZipArchive? _zip;
        private readonly ZipMemoryIndex? _memZip;
        private readonly ExcelReaderOptions _options;
        private readonly DecompressedByteCounter _decompressedBytes;
        private readonly (string Name, string Path, ExcelSheetVisibility Visibility)[] _sheets;
        private readonly bool[] _styleIsDate;
        private int _current;

        private byte[] _sharedFlat = [];
        private int[] _sharedOffsets = [0];
        private bool _sharedLoaded;
        private string?[]? _sharedStringCache;

        internal XlsxReader(Stream stream, bool leaveOpen, ExcelReaderOptions? options = null)
            : this(stream, leaveOpen, new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen: true), options)
        {
        }

        internal XlsxReader(Stream stream, bool leaveOpen, ZipArchive zip, ExcelReaderOptions? options = null)
        {
            _stream = stream;
            _leaveOpen = leaveOpen;
            _options = options ?? ExcelReaderOptions.Default;
            _decompressedBytes = new DecompressedByteCounter(_options.MaxTotalDecompressedBytes);
            _zip = zip;
            try
            {
                LimitChecks.ThrowIfTooManyEntries(_zip.Entries.Count, _options);
                using ZipPart wbPart = ZipEntryBytes.Read(_zip, "xl/workbook.xml", _decompressedBytes);
                using ZipPart relsPart = ZipEntryBytes.Read(_zip, "xl/_rels/workbook.xml.rels", _decompressedBytes);
                _sheets = ParseSheets(wbPart.Memory.Span, relsPart.Memory.Span);
                if (_sheets.Length == 0)
                {
                    throw new InvalidDataException("The workbook contains no sheets.");
                }
                using ZipPart stylesPart = ZipEntryBytes.Read(_zip, "xl/styles.xml", _decompressedBytes);
                _styleIsDate = ParseStyleDateFlags(stylesPart.Memory.Span);
                IsDate1904 = ParseDate1904(wbPart.Memory.Span);
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
            (string Name, string Path, ExcelSheetVisibility Visibility)[] sheets, bool[] styleIsDate, bool date1904,
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

        private XlsxReader(ZipMemoryIndex memZip,
            (string Name, string Path, ExcelSheetVisibility Visibility)[] sheets, bool[] styleIsDate, bool date1904,
            ExcelReaderOptions options, DecompressedByteCounter decompressedBytes)
        {
            _leaveOpen = true;
            _memZip = memZip;
            _options = options;
            _decompressedBytes = decompressedBytes;
            _sheets = sheets;
            _styleIsDate = styleIsDate;
            IsDate1904 = date1904;
        }

        internal static ValueTask<XlsxReader> CreateAsync(Stream stream, bool leaveOpen, ExcelReaderOptions? options = null, CancellationToken ct = default)
        {
            ExcelReaderOptions effectiveOptions = options ?? ExcelReaderOptions.Default;
            DecompressedByteCounter decompressedBytes = new(effectiveOptions.MaxTotalDecompressedBytes);
            return ZipReaderOpen.OpenAsync(stream, leaveOpen, effectiveOptions,
                zip => ParseAsync(stream, leaveOpen, zip, effectiveOptions, decompressedBytes, ct), ct);
        }

        internal static ValueTask<XlsxReader> CreateFromOpenZipAsync(
            Stream stream, bool leaveOpen, ZipArchive zip, ExcelReaderOptions? options, CancellationToken ct)
        {
            ExcelReaderOptions effectiveOptions = options ?? ExcelReaderOptions.Default;
            DecompressedByteCounter decompressedBytes = new(effectiveOptions.MaxTotalDecompressedBytes);
            return ZipReaderOpen.FromOpenZipAsync(stream, leaveOpen, zip, effectiveOptions,
                z => ParseAsync(stream, leaveOpen, z, effectiveOptions, decompressedBytes, ct));
        }

        private static async ValueTask<XlsxReader> ParseAsync(
            Stream stream, bool leaveOpen, ZipArchive zip, ExcelReaderOptions effectiveOptions,
            DecompressedByteCounter decompressedBytes, CancellationToken ct)
        {
            using ZipPart wbPart = await ZipEntryBytes.ReadAsync(zip, "xl/workbook.xml", decompressedBytes, ct).ConfigureAwait(false);
            using ZipPart relsPart = await ZipEntryBytes.ReadAsync(zip, "xl/_rels/workbook.xml.rels", decompressedBytes, ct).ConfigureAwait(false);
            var sheets = ParseSheets(wbPart.Memory.Span, relsPart.Memory.Span);
            if (sheets.Length == 0)
            {
                throw new InvalidDataException("The workbook contains no sheets.");
            }
            using ZipPart stylesPart = await ZipEntryBytes.ReadAsync(zip, "xl/styles.xml", decompressedBytes, ct).ConfigureAwait(false);
            var styleIsDate = ParseStyleDateFlags(stylesPart.Memory.Span);
            bool date1904 = ParseDate1904(wbPart.Memory.Span);
            return new XlsxReader(stream, leaveOpen, zip, sheets, styleIsDate, date1904, effectiveOptions, decompressedBytes);
        }

        /// <inheritdoc/>
        public string SheetName => _sheets[_current].Name;
        /// <inheritdoc/>
        public int SheetCount => _sheets.Length;
        /// <inheritdoc/>
        public string SheetNameAt(int index)
        {
            WorkbookLookups.ValidateSheetIndex(index, _sheets.Length);
            return _sheets[index].Name;
        }
        /// <inheritdoc/>
        public ExcelSheetVisibility SheetVisibility => _sheets[_current].Visibility;
        /// <inheritdoc/>
        public ExcelSheetVisibility SheetVisibilityAt(int index)
        {
            WorkbookLookups.ValidateSheetIndex(index, _sheets.Length);
            return _sheets[index].Visibility;
        }
        /// <inheritdoc/>
        public bool IsDate1904 { get; }

        internal bool IsDateStyle(int style)
        {
            return WorkbookLookups.IsDateStyle(_styleIsDate, style);
        }

        /// <inheritdoc/>
        public bool TryMoveToSheet(ReadOnlySpan<char> name)
        {
            if (!WorkbookLookups.TryFindSheetIndex(_sheets, name, static s => s.Name, out int index))
            {
                return false;
            }
            _current = index;
            return true;
        }

        /// <inheritdoc/>
        public void MoveToSheet(int index)
        {
            WorkbookLookups.ValidateSheetIndex(index, _sheets.Length);
            _current = index;
        }

        /// <inheritdoc/>
        public Enumerator GetEnumerator()
        {
            if (_memZip is not null)
            {
                EnsureSharedLoadedFromMemory();
                return GetEnumeratorFromMemory();
            }
            EnsureSharedLoaded();
            var entry = WorkbookLookups.GetWorksheetEntry(_zip!, _sheets[_current].Path);
            return new Enumerator(this, WorkbookLookups.OpenEntryStream(entry, _decompressedBytes, _options), entry.Length);
        }

        IExcelRowEnumerator IExcelRowReader<IExcelRowEnumerator>.GetEnumerator()
        {
            return GetEnumerator();
        }

        /// <inheritdoc/>
        public Enumerator GetAsyncEnumerator(CancellationToken ct = default)
        {
            if (_memZip is not null)
            {
                EnsureSharedLoadedFromMemory();
                return GetEnumeratorFromMemory();
            }
            return new Enumerator(this, WorkbookLookups.GetWorksheetEntry(_zip!, _sheets[_current].Path), ct);
        }

        IExcelRowEnumerator IExcelRowReader<IExcelRowEnumerator>.GetAsyncEnumerator(CancellationToken ct)
        {
            return GetAsyncEnumerator(ct);
        }

        internal ReadOnlySpan<byte> SharedSpan => _sharedFlat;

        internal string?[] SharedStringCache => _sharedStringCache ??= WorkbookLookups.CreateSharedStringCache(_sharedOffsets);

        /// <inheritdoc/>
        public void Dispose()
        {
            if (_sharedFlat.Length > 0)
            {
                ArrayPool<byte>.Shared.Return(_sharedFlat);
                _sharedFlat = [];
            }
            _memZip?.Dispose();
            _zip?.Dispose();
            if (!_leaveOpen)
            {
                _stream?.Dispose();
            }
        }

        /// <inheritdoc/>
        public async ValueTask DisposeAsync()
        {
            if (_sharedFlat.Length > 0)
            {
                ArrayPool<byte>.Shared.Return(_sharedFlat);
                _sharedFlat = [];
            }
            _memZip?.Dispose();
            if (_zip is not null)
            {
                await _zip.DisposeAsync().ConfigureAwait(false);
            }
            if (!_leaveOpen && _stream is not null)
            {
                await _stream.DisposeAsync().ConfigureAwait(false);
            }
        }
    }
}
