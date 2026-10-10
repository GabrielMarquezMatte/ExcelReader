using System.Buffers;
using System.Collections.ObjectModel;
using System.Diagnostics.CodeAnalysis;
using ExcelReader.Core.Reader.Internal;
using ExcelReader.Core.Reader.Sources;
using ExcelReader.Core.Reader.Zip;

namespace ExcelReader.Core.Reader.Xlsx
{
    /// <summary>An open Office Open XML (.xlsx) workbook. Its sheets are read independently, each streaming its cells without loading the whole file into memory.</summary>
    /// <remarks>See <see cref="IExcelWorkbook"/> for the threading and lifetime contract.</remarks>
    public sealed partial class XlsxWorkbook : IExcelWorkbook
    {
        private readonly ZipIndex _zip;
        private readonly ExcelReaderOptions _options;
        private readonly DecompressedByteCounter _decompressedBytes;
        private readonly (string Name, string Path, ExcelSheetVisibility Visibility)[] _sheets;
        private readonly bool[] _styleIsDate;
        private readonly ExcelSheetList<XlsxSheet> _sheetList;
        private readonly ReadOnlyCollection<ExcelTable> _tables;

        [SuppressMessage("Usage", "CA2213:Disposable fields should be disposed", Justification = "Disposed by ReleaseResources, which the lifetime runs after the last enumerator.")]
        private readonly OnceGate _sharedGate = new();
        private byte[] _sharedFlat = [];
        private int[] _sharedOffsets = [0];
        private string?[]? _sharedStringCache;

        internal ReaderLifetime Lifetime { get; }

        internal XlsxWorkbook(Stream stream, bool leaveOpen, ExcelReaderOptions? options = null)
            : this(ZipIndex.Create(ByteSource.FromStream(stream, leaveOpen), options ?? ExcelReaderOptions.Default), options ?? ExcelReaderOptions.Default)
        {
        }

        private XlsxWorkbook(ZipIndex zip, ExcelReaderOptions options)
        {
            Lifetime = new ReaderLifetime(ReleaseResources);
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
                _sheetList = CreateSheetList();
                _tables = WorkbookTables.Build(
                    TableDiscovery.Load(zip, _sheets, _decompressedBytes, XlsxTables.ParseTable), _sheetList);
            }
            catch
            {
                zip.Dispose();
                throw;
            }
        }

        private XlsxWorkbook(ZipIndex zip,
            (string Name, string Path, ExcelSheetVisibility Visibility)[] sheets, bool[] styleIsDate, bool date1904,
            List<TablePart> tables, ExcelReaderOptions options, DecompressedByteCounter decompressedBytes)
        {
            Lifetime = new ReaderLifetime(ReleaseResources);
            _zip = zip;
            _options = options;
            _decompressedBytes = decompressedBytes;
            _sheets = sheets;
            _styleIsDate = styleIsDate;
            IsDate1904 = date1904;
            _sheetList = CreateSheetList();
            _tables = WorkbookTables.Build(tables, _sheetList);
        }

        internal static XlsxWorkbook CreateFromMemory(ReadOnlyMemory<byte> data, ExcelReaderOptions? options = null)
        {
            ExcelReaderOptions effectiveOptions = options ?? ExcelReaderOptions.Default;
            return new XlsxWorkbook(ZipIndex.Create(data, effectiveOptions), effectiveOptions);
        }

        internal static XlsxWorkbook CreateFromIndex(ZipIndex zip, ExcelReaderOptions options)
        {
            return new XlsxWorkbook(zip, options);
        }

        internal static async ValueTask<XlsxWorkbook> CreateAsync(Stream stream, bool leaveOpen, ExcelReaderOptions? options = null, CancellationToken ct = default)
        {
            ExcelReaderOptions effectiveOptions = options ?? ExcelReaderOptions.Default;
            ByteSource source = await ByteSource.FromStreamAsync(stream, leaveOpen, ct).ConfigureAwait(false);
            ZipIndex zip = await ZipIndex.CreateAsync(source, effectiveOptions, ct).ConfigureAwait(false);
            return await CreateFromIndexAsync(zip, effectiveOptions, ct).ConfigureAwait(false);
        }

        internal static async ValueTask<XlsxWorkbook> CreateFromIndexAsync(ZipIndex zip, ExcelReaderOptions options, CancellationToken ct)
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
                List<TablePart> tables = await TableDiscovery.LoadAsync(zip, sheets, decompressedBytes, XlsxTables.ParseTable, ct).ConfigureAwait(false);
                return new XlsxWorkbook(zip, sheets, styleIsDate, date1904, tables, options, decompressedBytes);
            }
            catch
            {
                zip.Dispose();
                throw;
            }
        }

        /// <inheritdoc/>
        public int SheetCount => _sheets.Length;
        /// <inheritdoc/>
        public bool IsDate1904 { get; }

        internal bool IsDateStyle(int style)
        {
            return WorkbookLookups.IsDateStyle(_styleIsDate, style);
        }

        private ExcelSheetList<XlsxSheet> CreateSheetList()
        {
            XlsxSheet[] sheets = new XlsxSheet[_sheets.Length];
            for (int i = 0; i < sheets.Length; i++)
            {
                sheets[i] = new XlsxSheet(this, i, _sheets[i].Name, _sheets[i].Visibility);
            }
            return new ExcelSheetList<XlsxSheet>(sheets);
        }

        /// <summary>Gets the workbook's sheets, in workbook order. Reading the list opens nothing.</summary>
        /// <exception cref="ObjectDisposedException">The workbook was disposed.</exception>
        public ExcelSheetList<XlsxSheet> Sheets
        {
            get
            {
                Lifetime.ThrowIfClosed(this);
                return _sheetList;
            }
        }

        /// <summary>Finds a sheet by name, ignoring case. Opens nothing.</summary>
        /// <param name="name">The sheet name to look for.</param>
        /// <param name="sheet">The matching sheet, when one is found.</param>
        /// <returns><see langword="true"/> if a sheet with that name exists; otherwise <see langword="false"/>.</returns>
        /// <exception cref="ObjectDisposedException">The workbook was disposed.</exception>
        public bool TryGetSheet(ReadOnlySpan<char> name, out XlsxSheet sheet)
        {
            Lifetime.ThrowIfClosed(this);
            if (WorkbookLookups.TryFindSheetIndex(_sheets, name, static s => s.Name, out int index))
            {
                sheet = _sheetList[index];
                return true;
            }
            sheet = default;
            return false;
        }

        IExcelSheet IExcelWorkbook.SheetAt(int index)
        {
            return Sheets[index];
        }

        bool IExcelWorkbook.TryGetSheet(ReadOnlySpan<char> name, [MaybeNullWhen(false)] out IExcelSheet sheet)
        {
            bool found = TryGetSheet(name, out XlsxSheet typed);
            sheet = found ? typed : null;
            return found;
        }

        /// <summary>Gets the workbook's tables, ordered by sheet and then by the order they were created in. Opens nothing.</summary>
        /// <exception cref="ObjectDisposedException">The workbook was disposed.</exception>
        public IReadOnlyList<ExcelTable> Tables
        {
            get
            {
                Lifetime.ThrowIfClosed(this);
                return _tables;
            }
        }

        /// <summary>Finds a table by name, ignoring case. Opens nothing.</summary>
        /// <param name="name">The table name to look for.</param>
        /// <param name="table">The matching table, when one is found.</param>
        /// <returns><see langword="true"/> if a table with that name exists; otherwise <see langword="false"/>.</returns>
        /// <exception cref="ObjectDisposedException">The workbook was disposed.</exception>
        public bool TryGetTable(ReadOnlySpan<char> name, [MaybeNullWhen(false)] out ExcelTable table)
        {
            Lifetime.ThrowIfClosed(this);
            return WorkbookTables.TryFind(_tables, name, out table);
        }

        internal Enumerator OpenSheet(int index)
        {
            Lifetime.Acquire(this);
            try
            {
                EnsureSharedLoaded();
                ZipEntryRef entry = WorkbookLookups.GetWorksheetEntry(_zip, _sheets[index].Path);
                return new Enumerator(this, _zip.OpenEntryStream(entry, _decompressedBytes, _options), entry.UncompressedSize);
            }
            catch
            {
                Lifetime.Release();
                throw;
            }
        }

        internal Enumerator OpenSheetAsync(int index, CancellationToken ct)
        {
            if (_zip.HasMemory)
            {
                return OpenSheet(index);
            }
            Lifetime.Acquire(this);
            try
            {
                return new Enumerator(this, WorkbookLookups.GetWorksheetEntry(_zip, _sheets[index].Path), ct);
            }
            catch
            {
                Lifetime.Release();
                throw;
            }
        }

        /// <summary>Gets the workbook's first sheet: the same sheet as <c>Sheets[0]</c>. Opens nothing.</summary>
        /// <exception cref="ObjectDisposedException">The workbook was disposed.</exception>
        public XlsxSheet FirstSheet
        {
            get
            {
                return Sheets[0];
            }
        }

        IExcelSheet IExcelWorkbook.FirstSheet
        {
            get
            {
                return FirstSheet;
            }
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
            Lifetime.Close();
        }

        /// <inheritdoc/>
        public ValueTask DisposeAsync()
        {
            Lifetime.Close();
            return ValueTask.CompletedTask;
        }

        private void ReleaseResources()
        {
            ReturnSharedTable();
            _sharedGate.Dispose();
            _zip.Dispose();
        }

        private void ReturnSharedTable()
        {
            if (_sharedFlat.Length > 0)
            {
                ArrayPool<byte>.Shared.Return(_sharedFlat);
            }
            _sharedFlat = [];
            _sharedOffsets = [0];
        }
    }
}
