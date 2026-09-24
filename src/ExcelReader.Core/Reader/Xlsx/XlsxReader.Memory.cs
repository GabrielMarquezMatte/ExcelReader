using System.Buffers;
using ExcelReader.Core.Reader.Internal;
using ExcelReader.Core.Reader.Zip;

namespace ExcelReader.Core.Reader.Xlsx
{
    public sealed partial class XlsxReader
    {
        internal static XlsxReader CreateFromMemory(ReadOnlyMemory<byte> data, ExcelReaderOptions? options = null)
        {
            ExcelReaderOptions effectiveOptions = options ?? ExcelReaderOptions.Default;
            return CreateFromMemory(ZipMemoryIndex.Create(data, effectiveOptions), effectiveOptions);
        }

        internal static XlsxReader CreateFromMemory(ZipMemoryIndex memZip, ExcelReaderOptions effectiveOptions)
        {
            return ZipReaderOpen.FromMemory(memZip, zip => BuildFromMemory(zip, effectiveOptions));
        }

        private static XlsxReader BuildFromMemory(ZipMemoryIndex memZip, ExcelReaderOptions effectiveOptions)
        {
            DecompressedByteCounter decompressedBytes = new(effectiveOptions.MaxTotalDecompressedBytes);
            using ZipPart wbPart = memZip.OpenPartOrDefault("xl/workbook.xml"u8, decompressedBytes);
            using ZipPart relsPart = memZip.OpenPartOrDefault("xl/_rels/workbook.xml.rels"u8, decompressedBytes);
            (string Name, string Path, ExcelSheetVisibility Visibility)[] sheets = ParseSheets(wbPart.Memory.Span, relsPart.Memory.Span);
            if (sheets.Length == 0)
            {
                throw new InvalidDataException("The workbook contains no sheets.");
            }
            using ZipPart stylesPart = memZip.OpenPartOrDefault("xl/styles.xml"u8, decompressedBytes);
            bool[] styleIsDate = ParseStyleDateFlags(stylesPart.Memory.Span);
            bool date1904 = ParseDate1904(wbPart.Memory.Span);
            return new XlsxReader(memZip, sheets, styleIsDate, date1904, effectiveOptions, decompressedBytes);
        }

        private void EnsureSharedLoadedFromMemory()
        {
            if (_sharedLoaded)
            {
                return;
            }
            _sharedLoaded = true;
            if (!_memZip!.TryGetEntry("xl/sharedStrings.xml"u8, out ZipEntryRef entry))
            {
                return;
            }
            WorkbookLookups.ThrowIfSharedEntryTooLarge(entry.UncompressedSize, _decompressedBytes, _options);
            using ZipPart part = _memZip.OpenPart(entry, _decompressedBytes,
                nameof(ExcelReaderOptions.MaxSharedStringBytes), _options.MaxSharedStringBytes);
            ParseSharedFromMemory(part.Memory, entry.UncompressedSize);
        }

        private void ParseSharedFromMemory(ReadOnlyMemory<byte> content, long entryLength)
        {
            LimitChecks.ThrowIfEntryLengthExceeds(entryLength, Array.MaxLength, "ArrayMaxLength");
            int partLength = (int)entryLength;
            int growthCap = SharedFlatGrowthCap();
            var io = new BufferedStreamCursor(content, growthCap, nameof(ExcelReaderOptions.MaxSharedStringBytes));
            _sharedFlat = ArrayPool<byte>.Shared.Rent(Math.Max(1, partLength));
            _sharedOffsets = ParseSharedBody(io, stream: null, partLength);
        }

        private Enumerator GetEnumeratorFromMemory()
        {
            ZipEntryRef entry = WorkbookLookups.GetWorksheetEntry(_memZip!, _sheets[_current].Path);
            return new Enumerator(this, _memZip!.OpenEntryStream(entry, _decompressedBytes, _options), entry.UncompressedSize);
        }
    }
}
