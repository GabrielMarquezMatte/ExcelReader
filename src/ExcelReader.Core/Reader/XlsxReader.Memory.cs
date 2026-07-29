using System.Buffers;
using System.Diagnostics.CodeAnalysis;

namespace ExcelReader.Core.Reader
{
    // In-memory ZIP path: opens an XlsxReader directly over a
    // ReadOnlyMemory<byte> via ZipMemoryIndex instead of ZipArchive/Stream. No refills, no async
    // suspension — every part is already fully decompressed before the reader is constructed.
    public sealed partial class XlsxReader
    {
        [SuppressMessage("Reliability", "CA2000:Dispose objects before losing scope",
            Justification = "Ownership transfers to the (ZipMemoryIndex, ExcelReaderOptions) overload, which disposes it on failure and via the reader on success.")]
        internal static XlsxReader CreateFromMemory(ReadOnlyMemory<byte> data, ExcelReaderOptions? options = null)
        {
            ExcelReaderOptions effectiveOptions = options ?? ExcelReaderOptions.Default;
            return CreateFromMemory(ZipMemoryIndex.Create(data, effectiveOptions), effectiveOptions);
        }

        // Takes an already-built index (from Excel.Open's format peek) so the central directory isn't
        // walked a second time — the memory-path twin of CreateFromOpenZipAsync.
        internal static XlsxReader CreateFromMemory(ZipMemoryIndex memZip, ExcelReaderOptions effectiveOptions)
        {
            return ZipReaderOpen.FromMemory(memZip, zip => BuildFromMemory(zip, effectiveOptions));
        }

        private static XlsxReader BuildFromMemory(ZipMemoryIndex memZip, ExcelReaderOptions effectiveOptions)
        {
            DecompressedByteCounter decompressedBytes = new(effectiveOptions.MaxTotalDecompressedBytes);
            using ZipPart wbPart = memZip.OpenPartOrDefault("xl/workbook.xml"u8, decompressedBytes);
            using ZipPart relsPart = memZip.OpenPartOrDefault("xl/_rels/workbook.xml.rels"u8, decompressedBytes);
            (string Name, string Path)[] sheets = ParseSheets(wbPart.Memory.Span, relsPart.Memory.Span);
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

        // Reuses the streaming sst/si parser with a pre-filled, EOF-from-construction cursor: every
        // Fill it could call is unreachable (BufferedStreamCursor.Eof is already true), so passing a
        // null Stream is safe and the whole table is decoded in one pass with no growth loop.
        private void ParseSharedFromMemory(ReadOnlyMemory<byte> content, long entryLength)
        {
            LimitChecks.ThrowIfEntryLengthExceeds(entryLength, Array.MaxLength, "ArrayMaxLength");
            int partLength = (int)entryLength;
            int growthCap = SharedFlatGrowthCap();
            var io = new BufferedStreamCursor(content, growthCap, nameof(ExcelReaderOptions.MaxSharedStringBytes));
            _sharedFlat = ArrayPool<byte>.Shared.Rent(Math.Max(1, partLength));
            _sharedOffsets = ParseSharedBody(io, stream: null, partLength);
        }

        // Worksheet entry only: opens a Stream (DeflateStream, optionally wrapped in PrefetchStream via
        // ZipMemoryIndex.OpenEntryStream) instead of eagerly materializing a ZipPart, so
        // PrefetchDecompression overlaps inflate with row parsing on this path exactly as it does for
        // the ZipArchive-backed reader.
        private Enumerator GetEnumeratorFromMemory()
        {
            ZipEntryRef entry = WorkbookLookups.GetWorksheetEntry(_memZip!, _sheets, _current);
            return new Enumerator(this, _memZip!.OpenEntryStream(entry, _decompressedBytes, _options), entry.UncompressedSize);
        }
    }
}
