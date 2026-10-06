using System.Buffers;
using ExcelReader.Core.Reader.Internal;
using ExcelReader.Core.Reader.Sources;
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

        private readonly ZipIndex? _zip;
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
            : this(ZipIndex.Create(ByteSource.FromStream(stream, leaveOpen), options ?? ExcelReaderOptions.Default), options ?? ExcelReaderOptions.Default)
        {
        }

        private XlsbReader(ZipIndex zip, ExcelReaderOptions options)
        {
            _zip = zip;
            _options = options;
            _decompressedBytes = new DecompressedByteCounter(options.MaxTotalDecompressedBytes);
            try
            {
                using ZipPart wbPart = zip.OpenPartOrDefault("xl/workbook.bin"u8, _decompressedBytes);
                using ZipPart relsPart = zip.OpenPartOrDefault("xl/_rels/workbook.bin.rels"u8, _decompressedBytes);
                _sheets = XlsbWorkbook.ParseSheets(wbPart.Memory.Span, relsPart.Memory.Span);
                if (_sheets.Length == 0)
                {
                    throw new InvalidDataException("The workbook contains no sheets.");
                }
                using ZipPart stylesPart = zip.OpenPartOrDefault("xl/styles.bin"u8, _decompressedBytes);
                _styleIsDate = XlsbStyles.ParseStyleDateFlags(stylesPart.Memory.Span);
                IsDate1904 = XlsbWorkbook.ParseDate1904(wbPart.Memory.Span);
                (_sharedFlat, _sharedOffsets, _pooledSharedFlat) = LoadSharedStrings(zip, _decompressedBytes, options);
            }
            catch
            {
                zip.Dispose();
                throw;
            }
        }

        private XlsbReader(ZipIndex zip,
            (string Name, string Path, ExcelSheetVisibility Visibility)[] sheets, bool[] styleIsDate, bool date1904,
            byte[] sharedFlat, int[] sharedOffsets, bool pooledSharedFlat, ExcelReaderOptions options, DecompressedByteCounter decompressedBytes)
        {
            _zip = zip;
            _options = options;
            _decompressedBytes = decompressedBytes;
            _sheets = sheets;
            _styleIsDate = styleIsDate;
            IsDate1904 = date1904;
            _sharedFlat = sharedFlat;
            _sharedOffsets = sharedOffsets;
            _pooledSharedFlat = pooledSharedFlat;
        }

        internal static XlsbReader CreateFromMemory(ReadOnlyMemory<byte> data, ExcelReaderOptions? options = null)
        {
            ExcelReaderOptions effectiveOptions = options ?? ExcelReaderOptions.Default;
            return new XlsbReader(ZipIndex.Create(data, effectiveOptions), effectiveOptions);
        }

        internal static XlsbReader CreateFromIndex(ZipIndex zip, ExcelReaderOptions options)
        {
            return new XlsbReader(zip, options);
        }

        internal static async ValueTask<XlsbReader> CreateAsync(Stream stream, bool leaveOpen, ExcelReaderOptions? options = null, CancellationToken ct = default)
        {
            ExcelReaderOptions effectiveOptions = options ?? ExcelReaderOptions.Default;
            ByteSource source = await ByteSource.FromStreamAsync(stream, leaveOpen, ct).ConfigureAwait(false);
            ZipIndex zip = await ZipIndex.CreateAsync(source, effectiveOptions, ct).ConfigureAwait(false);
            return await CreateFromIndexAsync(zip, effectiveOptions, ct).ConfigureAwait(false);
        }

        internal static async ValueTask<XlsbReader> CreateFromIndexAsync(ZipIndex zip, ExcelReaderOptions options, CancellationToken ct)
        {
            try
            {
                DecompressedByteCounter decompressedBytes = new(options.MaxTotalDecompressedBytes);
                using ZipPart wbPart = await zip.OpenPartOrDefaultAsync("xl/workbook.bin"u8, decompressedBytes, ct).ConfigureAwait(false);
                using ZipPart relsPart = await zip.OpenPartOrDefaultAsync("xl/_rels/workbook.bin.rels"u8, decompressedBytes, ct).ConfigureAwait(false);
                (string Name, string Path, ExcelSheetVisibility Visibility)[] sheets = XlsbWorkbook.ParseSheets(wbPart.Memory.Span, relsPart.Memory.Span);
                if (sheets.Length == 0)
                {
                    throw new InvalidDataException("The workbook contains no sheets.");
                }
                using ZipPart stylesPart = await zip.OpenPartOrDefaultAsync("xl/styles.bin"u8, decompressedBytes, ct).ConfigureAwait(false);
                bool[] styleIsDate = XlsbStyles.ParseStyleDateFlags(stylesPart.Memory.Span);
                bool date1904 = XlsbWorkbook.ParseDate1904(wbPart.Memory.Span);
                (byte[] flat, int[] offsets, bool pooled) = await LoadSharedStringsAsync(zip, decompressedBytes, options, ct).ConfigureAwait(false);
                return new XlsbReader(zip, sheets, styleIsDate, date1904, flat, offsets, pooled, options, decompressedBytes);
            }
            catch
            {
                zip.Dispose();
                throw;
            }
        }

        private static (byte[] Flat, int[] Offsets, bool Pooled) LoadSharedStrings(
            ZipIndex zip, DecompressedByteCounter decompressedBytes, ExcelReaderOptions options)
        {
            if (!zip.TryGetEntry("xl/sharedStrings.bin"u8, out ZipEntryRef entry))
            {
                return ([], [0], false);
            }
            WorkbookLookups.ThrowIfSharedEntryTooLarge(entry.UncompressedSize, decompressedBytes, options);
            if (zip.HasMemory)
            {
                using ZipPart part = zip.OpenPart(entry, decompressedBytes,
                    nameof(ExcelReaderOptions.MaxSharedStringBytes), options.MaxSharedStringBytes);
                (byte[] parsedFlat, int[] parsedOffsets) = XlsbSharedStrings.Parse(part.Memory.Span, options);
                return (parsedFlat, parsedOffsets, false);
            }
            using LimitedReadStream stream = zip.OpenEntryStream(entry, decompressedBytes, options,
                nameof(ExcelReaderOptions.MaxSharedStringBytes), options.MaxSharedStringBytes);
            (byte[] flat, int[] offsets) = XlsbSharedStrings.ParseStreaming(stream, entry.UncompressedSize, options);
            return (flat, offsets, flat.Length != 0);
        }

        private static async ValueTask<(byte[] Flat, int[] Offsets, bool Pooled)> LoadSharedStringsAsync(
            ZipIndex zip, DecompressedByteCounter decompressedBytes, ExcelReaderOptions options, CancellationToken ct)
        {
            if (zip.HasMemory || !zip.TryGetEntry("xl/sharedStrings.bin"u8, out ZipEntryRef entry))
            {
                return LoadSharedStrings(zip, decompressedBytes, options);
            }
            WorkbookLookups.ThrowIfSharedEntryTooLarge(entry.UncompressedSize, decompressedBytes, options);
            LimitedReadStream stream = await zip.OpenEntryStreamAsync(entry, decompressedBytes, options, ct,
                nameof(ExcelReaderOptions.MaxSharedStringBytes), options.MaxSharedStringBytes).ConfigureAwait(false);
            await using (stream.ConfigureAwait(false))
            {
                (byte[] flat, int[] offsets) = await XlsbSharedStrings.ParseStreamingAsync(stream, entry.UncompressedSize, options, ct).ConfigureAwait(false);
                return (flat, offsets, flat.Length != 0);
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
            ZipEntryRef entry = WorkbookLookups.GetWorksheetEntry(_zip!, _sheets![_current].Path);
            return new Enumerator(this, _zip!.OpenEntryStream(entry, _decompressedBytes, _options), entry.UncompressedSize);
        }

        IExcelRowEnumerator IExcelRowReader<IExcelRowEnumerator>.GetEnumerator()
        {
            return GetEnumerator();
        }

        /// <inheritdoc/>
        public Enumerator GetAsyncEnumerator(CancellationToken ct = default)
        {
            if (_zip!.HasMemory)
            {
                return GetEnumerator();
            }
            return new Enumerator(this, WorkbookLookups.GetWorksheetEntry(_zip, _sheets![_current].Path), ct);
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
            _zip?.Dispose();
        }

        /// <inheritdoc/>
        public ValueTask DisposeAsync()
        {
            Dispose();
            return ValueTask.CompletedTask;
        }

    }
}
