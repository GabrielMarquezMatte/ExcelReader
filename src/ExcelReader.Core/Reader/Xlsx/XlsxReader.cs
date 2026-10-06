using System.Buffers;
using ExcelReader.Core.Reader.Internal;
using ExcelReader.Core.Reader.Sources;
using ExcelReader.Core.Reader.Zip;

namespace ExcelReader.Core.Reader.Xlsx
{
    /// <summary>Reads rows from an Office Open XML (.xlsx) workbook, streaming each sheet's cells without loading the whole file into memory.</summary>
    public sealed partial class XlsxReader : IExcelRowReader, IExcelRowReader<XlsxReader.Enumerator>
    {
        private readonly ZipIndex _zip;
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
            : this(ZipIndex.Create(ByteSource.FromStream(stream, leaveOpen), options ?? ExcelReaderOptions.Default), options ?? ExcelReaderOptions.Default)
        {
        }

        private XlsxReader(ZipIndex zip, ExcelReaderOptions options)
        {
            _zip = zip;
            _options = options;
            _decompressedBytes = new DecompressedByteCounter(options.MaxTotalDecompressedBytes);
            try
            {
                using ZipPart wbPart = zip.OpenPartOrDefault("xl/workbook.xml"u8, _decompressedBytes);
                using ZipPart relsPart = zip.OpenPartOrDefault("xl/_rels/workbook.xml.rels"u8, _decompressedBytes);
                _sheets = ParseSheets(wbPart.Memory.Span, relsPart.Memory.Span);
                if (_sheets.Length == 0)
                {
                    throw new InvalidDataException("The workbook contains no sheets.");
                }
                using ZipPart stylesPart = zip.OpenPartOrDefault("xl/styles.xml"u8, _decompressedBytes);
                _styleIsDate = ParseStyleDateFlags(stylesPart.Memory.Span);
                IsDate1904 = ParseDate1904(wbPart.Memory.Span);
            }
            catch
            {
                zip.Dispose();
                throw;
            }
        }

        private XlsxReader(ZipIndex zip,
            (string Name, string Path, ExcelSheetVisibility Visibility)[] sheets, bool[] styleIsDate, bool date1904,
            ExcelReaderOptions options, DecompressedByteCounter decompressedBytes)
        {
            _zip = zip;
            _options = options;
            _decompressedBytes = decompressedBytes;
            _sheets = sheets;
            _styleIsDate = styleIsDate;
            IsDate1904 = date1904;
        }

        internal static XlsxReader CreateFromMemory(ReadOnlyMemory<byte> data, ExcelReaderOptions? options = null)
        {
            ExcelReaderOptions effectiveOptions = options ?? ExcelReaderOptions.Default;
            return new XlsxReader(ZipIndex.Create(data, effectiveOptions), effectiveOptions);
        }

        internal static XlsxReader CreateFromIndex(ZipIndex zip, ExcelReaderOptions options)
        {
            return new XlsxReader(zip, options);
        }

        internal static async ValueTask<XlsxReader> CreateAsync(Stream stream, bool leaveOpen, ExcelReaderOptions? options = null, CancellationToken ct = default)
        {
            ExcelReaderOptions effectiveOptions = options ?? ExcelReaderOptions.Default;
            ByteSource source = await ByteSource.FromStreamAsync(stream, leaveOpen, ct).ConfigureAwait(false);
            ZipIndex zip = await ZipIndex.CreateAsync(source, effectiveOptions, ct).ConfigureAwait(false);
            return await CreateFromIndexAsync(zip, effectiveOptions, ct).ConfigureAwait(false);
        }

        internal static async ValueTask<XlsxReader> CreateFromIndexAsync(ZipIndex zip, ExcelReaderOptions options, CancellationToken ct)
        {
            try
            {
                DecompressedByteCounter decompressedBytes = new(options.MaxTotalDecompressedBytes);
                using ZipPart wbPart = await zip.OpenPartOrDefaultAsync("xl/workbook.xml"u8, decompressedBytes, ct).ConfigureAwait(false);
                using ZipPart relsPart = await zip.OpenPartOrDefaultAsync("xl/_rels/workbook.xml.rels"u8, decompressedBytes, ct).ConfigureAwait(false);
                (string Name, string Path, ExcelSheetVisibility Visibility)[] sheets = ParseSheets(wbPart.Memory.Span, relsPart.Memory.Span);
                if (sheets.Length == 0)
                {
                    throw new InvalidDataException("The workbook contains no sheets.");
                }
                using ZipPart stylesPart = await zip.OpenPartOrDefaultAsync("xl/styles.xml"u8, decompressedBytes, ct).ConfigureAwait(false);
                bool[] styleIsDate = ParseStyleDateFlags(stylesPart.Memory.Span);
                bool date1904 = ParseDate1904(wbPart.Memory.Span);
                return new XlsxReader(zip, sheets, styleIsDate, date1904, options, decompressedBytes);
            }
            catch
            {
                zip.Dispose();
                throw;
            }
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
            EnsureSharedLoaded();
            ZipEntryRef entry = WorkbookLookups.GetWorksheetEntry(_zip, _sheets[_current].Path);
            return new Enumerator(this, _zip.OpenEntryStream(entry, _decompressedBytes, _options), entry.UncompressedSize);
        }

        IExcelRowEnumerator IExcelRowReader<IExcelRowEnumerator>.GetEnumerator()
        {
            return GetEnumerator();
        }

        /// <inheritdoc/>
        public Enumerator GetAsyncEnumerator(CancellationToken ct = default)
        {
            if (_zip.HasMemory)
            {
                return GetEnumerator();
            }
            return new Enumerator(this, WorkbookLookups.GetWorksheetEntry(_zip, _sheets[_current].Path), ct);
        }

        IExcelRowEnumerator IExcelRowReader<IExcelRowEnumerator>.GetAsyncEnumerator(CancellationToken ct)
        {
            return GetAsyncEnumerator(ct);
        }

        internal ReadOnlySpan<byte> SharedSpan => _sharedFlat;

        internal string?[] SharedStringCache => Volatile.Read(ref _sharedStringCache) ?? CreateSharedStringCache();

        private string?[] CreateSharedStringCache()
        {
            string?[] created = WorkbookLookups.CreateSharedStringCache(_sharedOffsets);
            return Interlocked.CompareExchange(ref _sharedStringCache, created, null) ?? created;
        }

        /// <inheritdoc/>
        public void Dispose()
        {
            if (_sharedFlat.Length > 0)
            {
                ArrayPool<byte>.Shared.Return(_sharedFlat);
                _sharedFlat = [];
            }
            _zip.Dispose();
        }

        /// <inheritdoc/>
        public ValueTask DisposeAsync()
        {
            Dispose();
            return ValueTask.CompletedTask;
        }
    }
}
