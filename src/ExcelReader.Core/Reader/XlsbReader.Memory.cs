using System.Diagnostics.CodeAnalysis;

namespace ExcelReader.Core.Reader
{
    // In-memory ZIP path: opens an XlsbReader directly over a
    // ReadOnlyMemory<byte> via ZipMemoryIndex instead of ZipArchive/Stream. No refills, no async
    // suspension — every part is already fully decompressed before the reader is constructed.
    public sealed partial class XlsbReader
    {
        [SuppressMessage("Reliability", "CA2000:Dispose objects before losing scope",
            Justification = "Ownership transfers to the (ZipMemoryIndex, ExcelReaderOptions) overload, which disposes it on failure and via the reader on success.")]
        internal static XlsbReader CreateFromMemory(ReadOnlyMemory<byte> data, ExcelReaderOptions? options = null)
        {
            ExcelReaderOptions effectiveOptions = options ?? ExcelReaderOptions.Default;
            return CreateFromMemory(ZipMemoryIndex.Create(data, effectiveOptions), effectiveOptions);
        }

        // Takes an already-built index (from Excel.Open's format peek) so the central directory isn't
        // walked a second time — the memory-path twin of CreateFromOpenZipAsync.
        internal static XlsbReader CreateFromMemory(ZipMemoryIndex memZip, ExcelReaderOptions effectiveOptions)
        {
            return ZipReaderOpen.FromMemory(memZip, zip => BuildFromMemory(zip, effectiveOptions));
        }

        private static XlsbReader BuildFromMemory(ZipMemoryIndex memZip, ExcelReaderOptions effectiveOptions)
        {
            DecompressedByteCounter decompressedBytes = new(effectiveOptions.MaxTotalDecompressedBytes);
            using ZipPart wbPart = memZip.OpenPartOrDefault("xl/workbook.bin"u8, decompressedBytes);
            using ZipPart relsPart = memZip.OpenPartOrDefault("xl/_rels/workbook.bin.rels"u8, decompressedBytes);
            (string Name, string Path)[] sheets = XlsbWorkbook.ParseSheets(wbPart.Memory.Span, relsPart.Memory.Span);
            if (sheets.Length == 0)
            {
                throw new InvalidDataException("The workbook contains no sheets.");
            }
            using ZipPart stylesPart = memZip.OpenPartOrDefault("xl/styles.bin"u8, decompressedBytes);
            bool[] styleIsDate = XlsbStyles.ParseStyleDateFlags(stylesPart.Memory.Span);
            bool date1904 = XlsbWorkbook.ParseDate1904(wbPart.Memory.Span);
            (byte[] flat, int[] offsets) = LoadSharedStringsFromMemory(memZip, decompressedBytes, effectiveOptions);
            return new XlsbReader(memZip, sheets, styleIsDate, date1904, flat, offsets, effectiveOptions, decompressedBytes);
        }

        private static (byte[] Flat, int[] Offsets) LoadSharedStringsFromMemory(
            ZipMemoryIndex memZip, DecompressedByteCounter decompressedBytes, ExcelReaderOptions options)
        {
            if (!memZip.TryGetEntry("xl/sharedStrings.bin"u8, out ZipEntryRef entry))
            {
                return ([], [0]);
            }
            WorkbookLookups.ThrowIfSharedEntryTooLarge(entry.UncompressedSize, decompressedBytes, options);
            using ZipPart part = memZip.OpenPart(entry, decompressedBytes,
                nameof(ExcelReaderOptions.MaxSharedStringBytes), options.MaxSharedStringBytes);
            return XlsbSharedStrings.Parse(part.Memory.Span, options);
        }

        // Worksheet entry only: opens a Stream (DeflateStream, optionally wrapped in PrefetchStream via
        // ZipMemoryIndex.OpenEntryStream) instead of eagerly materializing a ZipPart, so
        // PrefetchDecompression overlaps inflate with row parsing on this path exactly as it does for
        // the ZipArchive-backed reader.
        private Enumerator GetEnumeratorFromMemory()
        {
            ZipEntryRef entry = WorkbookLookups.GetWorksheetEntry(_memZip!, _sheets!, _current);
            return new Enumerator(this, _memZip!.OpenEntryStream(entry, _decompressedBytes, _options), entry.UncompressedSize);
        }
    }
}
