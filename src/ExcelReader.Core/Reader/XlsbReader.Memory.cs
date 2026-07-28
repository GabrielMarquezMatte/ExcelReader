using System.Diagnostics.CodeAnalysis;

namespace ExcelReader.Core.Reader
{
    // In-memory ZIP path (docs/in-memory-zip.md, Z4): opens an XlsbReader directly over a
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
        // walked a second time — the memory-path twin of CreateFromOpenZipAsync. Owns dispose-on-failure
        // either way: on success memZip's lifetime transfers to the returned reader.
        [SuppressMessage("IDisposableAnalyzers.Correctness", "IDISP007:Don't dispose injected",
            Justification = "memZip's lifetime transfers to this call; disposing it here on failure is correct ownership, not disposing a borrowed dependency.")]
        internal static XlsbReader CreateFromMemory(ZipMemoryIndex memZip, ExcelReaderOptions effectiveOptions)
        {
            try
            {
                return BuildFromMemory(memZip, effectiveOptions);
            }
            catch
            {
                memZip.Dispose();
                throw;
            }
        }

        private static XlsbReader BuildFromMemory(ZipMemoryIndex memZip, ExcelReaderOptions effectiveOptions)
        {
            DecompressedByteCounter decompressedBytes = new(effectiveOptions.MaxTotalDecompressedBytes);
            using ZipPart wbPart = OpenPartOrDefault(memZip, "xl/workbook.bin"u8, decompressedBytes);
            using ZipPart relsPart = OpenPartOrDefault(memZip, "xl/_rels/workbook.bin.rels"u8, decompressedBytes);
            (string Name, string Path)[] sheets = XlsbWorkbook.ParseSheets(wbPart.Memory.Span, relsPart.Memory.Span);
            if (sheets.Length == 0)
            {
                throw new InvalidDataException("The workbook contains no sheets.");
            }
            using ZipPart stylesPart = OpenPartOrDefault(memZip, "xl/styles.bin"u8, decompressedBytes);
            bool[] styleIsDate = XlsbStyles.ParseStyleDateFlags(stylesPart.Memory.Span);
            bool date1904 = XlsbWorkbook.ParseDate1904(wbPart.Memory.Span);
            (byte[] flat, int[] offsets) = LoadSharedStringsFromMemory(memZip, decompressedBytes, effectiveOptions);
            return new XlsbReader(memZip, sheets, styleIsDate, date1904, flat, offsets, effectiveOptions, decompressedBytes);
        }

        // default(ZipPart) (empty Memory, nothing to return on Dispose) stands in for a missing part,
        // mirroring ZipEntryBytes.Read's "return [] when the entry is absent" behavior on the streamed path.
        private static ZipPart OpenPartOrDefault(ZipMemoryIndex memZip, ReadOnlySpan<byte> utf8Name, DecompressedByteCounter counter)
        {
            return memZip.TryGetEntry(utf8Name, out ZipEntryRef entry) ? memZip.OpenPart(entry, counter) : default;
        }

        private static (byte[] Flat, int[] Offsets) LoadSharedStringsFromMemory(
            ZipMemoryIndex memZip, DecompressedByteCounter decompressedBytes, ExcelReaderOptions options)
        {
            if (!memZip.TryGetEntry("xl/sharedStrings.bin"u8, out ZipEntryRef entry))
            {
                return ([], [0]);
            }
            ThrowIfSharedEntryTooLarge(entry.UncompressedSize, decompressedBytes, options);
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
