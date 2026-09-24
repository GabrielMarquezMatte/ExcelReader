using System.Buffers;
using System.IO.Compression;
using ExcelReader.Core.Reader.Internal;
using ExcelReader.Core.Reader.Zip;

namespace ExcelReader.Core.Reader.Xlsb
{
    /// <summary>Reads rows from a binary Excel (.xlsb / BIFF12) workbook, streaming each sheet's cells without loading the whole file into memory.</summary>
    /// <remarks>
    /// Uses the same ZIP/OPC container as .xlsx, but worksheet parts are binary BIFF12 records. The workbook,
    /// styles, and shared-string parts are read once at open time (they're small); worksheets are streamed on
    /// demand by the enumerator.
    /// </remarks>
    public sealed partial class XlsbReader : IExcelRowReader, IExcelRowReader<XlsbReader.Enumerator>
    {
        private readonly byte[] _sharedFlat = [];
        private readonly int[] _sharedOffsets = [0];
        private readonly bool _pooledSharedFlat;
        private string?[]? _sharedStringCache;
        private readonly bool[] _styleIsDate = [];
        private readonly ExcelReaderOptions _options;
        private readonly DecompressedByteCounter _decompressedBytes;

        private readonly ZipArchive? _zip;
        private readonly ZipMemoryIndex? _memZip;
        private readonly Stream? _stream;
        private readonly bool _leaveOpen;
        private readonly (string Name, string Path, ExcelSheetVisibility Visibility)[]? _sheets;
        private int _current;
        private int _disposed;

        internal XlsbReader(byte[] sharedFlat, int[] sharedOffsets, bool[] styleIsDate, bool date1904)
        {
            _options = ExcelReaderOptions.Default;
            _decompressedBytes = new DecompressedByteCounter(_options.MaxTotalDecompressedBytes);
            _sharedFlat = sharedFlat;
            _sharedOffsets = sharedOffsets;
            _styleIsDate = styleIsDate;
            IsDate1904 = date1904;
        }

        internal XlsbReader(Stream stream, bool leaveOpen, ExcelReaderOptions? options = null)
            : this(stream, leaveOpen, ZipReaderOpen.Open(stream, leaveOpen), options)
        {
        }

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
                using ZipPart wbPart = ZipEntryBytes.Read(_zip, "xl/workbook.bin", _decompressedBytes);
                using ZipPart relsPart = ZipEntryBytes.Read(_zip, "xl/_rels/workbook.bin.rels", _decompressedBytes);
                _sheets = XlsbWorkbook.ParseSheets(wbPart.Memory.Span, relsPart.Memory.Span);
                if (_sheets.Length == 0)
                {
                    throw new InvalidDataException("The workbook contains no sheets.");
                }
                using ZipPart stylesPart = ZipEntryBytes.Read(_zip, "xl/styles.bin", _decompressedBytes);
                _styleIsDate = XlsbStyles.ParseStyleDateFlags(stylesPart.Memory.Span);
                IsDate1904 = XlsbWorkbook.ParseDate1904(wbPart.Memory.Span);
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

        private XlsbReader(Stream stream, bool leaveOpen, ZipArchive zip,
            (string Name, string Path, ExcelSheetVisibility Visibility)[] sheets, bool[] styleIsDate, bool date1904,
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

        private XlsbReader(ZipMemoryIndex memZip,
            (string Name, string Path, ExcelSheetVisibility Visibility)[] sheets, bool[] styleIsDate, bool date1904,
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
            using ZipPart wbPart = await ZipEntryBytes.ReadAsync(zip, "xl/workbook.bin", decompressedBytes, ct).ConfigureAwait(false);
            using ZipPart relsPart = await ZipEntryBytes.ReadAsync(zip, "xl/_rels/workbook.bin.rels", decompressedBytes, ct).ConfigureAwait(false);
            var sheets = XlsbWorkbook.ParseSheets(wbPart.Memory.Span, relsPart.Memory.Span);
            if (sheets.Length == 0)
            {
                throw new InvalidDataException("The workbook contains no sheets.");
            }
            using ZipPart stylesPart = await ZipEntryBytes.ReadAsync(zip, "xl/styles.bin", decompressedBytes, ct).ConfigureAwait(false);
            var styleIsDate = XlsbStyles.ParseStyleDateFlags(stylesPart.Memory.Span);
            bool date1904 = XlsbWorkbook.ParseDate1904(wbPart.Memory.Span);
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


        /// <inheritdoc/>
        public bool IsDate1904 { get; }
        /// <inheritdoc/>
        public string SheetName => _sheets![_current].Name;
        /// <inheritdoc/>
        public int SheetCount => _sheets!.Length;
        /// <inheritdoc/>
        public string SheetNameAt(int index)
        {
            WorkbookLookups.ValidateSheetIndex(index, _sheets!.Length);
            return _sheets[index].Name;
        }
        /// <inheritdoc/>
        public ExcelSheetVisibility SheetVisibility => _sheets![_current].Visibility;
        /// <inheritdoc/>
        public ExcelSheetVisibility SheetVisibilityAt(int index)
        {
            WorkbookLookups.ValidateSheetIndex(index, _sheets!.Length);
            return _sheets[index].Visibility;
        }

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


        internal ReadOnlySpan<byte> SharedSpan => _sharedFlat;

        internal string?[] SharedStringCache => _sharedStringCache ??= WorkbookLookups.CreateSharedStringCache(_sharedOffsets);

        internal bool IsDateStyle(int style)
        {
            return WorkbookLookups.IsDateStyle(_styleIsDate, style);
        }

        /// <inheritdoc/>
        public Enumerator GetEnumerator()
        {
            if (_memZip is not null)
            {
                return GetEnumeratorFromMemory();
            }
            var entry = WorkbookLookups.GetWorksheetEntry(_zip!, _sheets![_current].Path);
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
                return GetEnumeratorFromMemory();
            }
            return new Enumerator(this, WorkbookLookups.GetWorksheetEntry(_zip!, _sheets![_current].Path), ct);
        }

        IExcelRowEnumerator IExcelRowReader<IExcelRowEnumerator>.GetAsyncEnumerator(CancellationToken ct)
        {
            return GetAsyncEnumerator(ct);
        }


        /// <inheritdoc/>
        public void Dispose()
        {
            if (Interlocked.Exchange(ref _disposed, 1) != 0)
            {
                return;
            }
            if (_pooledSharedFlat)
            {
                ArrayPool<byte>.Shared.Return(_sharedFlat);
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
            if (Interlocked.Exchange(ref _disposed, 1) != 0)
            {
                return;
            }
            if (_pooledSharedFlat)
            {
                ArrayPool<byte>.Shared.Return(_sharedFlat);
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
