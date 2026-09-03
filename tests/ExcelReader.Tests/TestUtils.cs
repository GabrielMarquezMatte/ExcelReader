using System.Buffers;
using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.IO.Compression;
using System.Text;
using ExcelReader.Core.Enums;
using ExcelReader.Core.Reader;
using ExcelReader.Core.Writer;

namespace ExcelReader.Tests
{
    // Shared seeded-mutation fuzz harness used by FuzzTests, MemoryZipParityTests, and
    // ZipMemoryIndexTests. Previously each file had its own copy of MutateCopy and its own
    // AcceptableExceptionTypes list, and the lists had already drifted apart from each other:
    // Span-bounds violation surfacing as ArgumentOutOfRangeException, or an unchecked-arithmetic
    // wrap surfacing as OverflowException, is the same "parser forgot a check" bug class as
    // IndexOutOfRangeException — which none of the three lists ever accepted. Only exceptions
    // that mean "this input was deliberately and correctly rejected" belong here.
    internal static class FuzzMutation
    {
        internal static readonly Type[] AcceptableExceptionTypes =
        [
            typeof(InvalidDataException),
            typeof(ExcelLimitExceededException),
            typeof(ExcelEncryptionException),
            typeof(IOException),
            typeof(NotSupportedException),
            typeof(FormatException),
        ];

        internal static bool IsAcceptable(Exception ex)
        {
            foreach (Type acceptableType in AcceptableExceptionTypes)
            {
                if (acceptableType.IsInstanceOfType(ex))
                {
                    return true;
                }
            }
            return false;
        }

        // Flips 1-8 random bytes to random values; a copy so the seed itself is never mutated.
        // Deliberately not cryptographically secure — CA5394 doesn't apply: a seeded, reproducible
        // PRNG is exactly what makes a fuzz failure pinpoint-able, unlike a CSPRNG would be.
        [SuppressMessage("Security", "CA5394:Do not use insecure randomness",
            Justification = "Fuzzing needs a reproducible seeded PRNG, not cryptographic randomness.")]
        internal static byte[] MutateCopy(byte[] seed, Random rng, out int[] positions)
        {
            byte[] copy = (byte[])seed.Clone();
            int count = rng.Next(1, 9);
            positions = new int[count];
            for (int i = 0; i < count; i++)
            {
                int pos = rng.Next(copy.Length);
                positions[i] = pos;
                copy[pos] = (byte)rng.Next(256);
            }
            return copy;
        }

        // An attacker-controlled count driving an allocation the configured limits never see
        // doesn't necessarily throw at all — it can succeed while
        // burning far more time/memory than a few-KB seed file could ever legitimately need. A round
        // that completes "successfully" after allocating hundreds of MB, or after seconds of spinning,
        // is not a graceful rejection; it's the amplification attack working. The caps here are set
        // generously above anything these tiny seeds should ever legitimately need (they normally
        // allocate low hundreds of KB and run in low single-digit milliseconds), so tripping this is a
        // real finding, not noise.
        private const long MaxAllocatedBytesPerRound = 32L * 1024 * 1024;
        private static readonly TimeSpan MaxDurationPerRound = TimeSpan.FromSeconds(2);

        internal static void RunBounded(Action action)
        {
            long before = GC.GetAllocatedBytesForCurrentThread();
            var stopwatch = Stopwatch.StartNew();
            action();
            stopwatch.Stop();
            long allocated = GC.GetAllocatedBytesForCurrentThread() - before;
            if (allocated > MaxAllocatedBytesPerRound)
            {
                throw new InvalidOperationException(
                    string.Create(CultureInfo.InvariantCulture,
                        $"Round allocated {allocated:N0} bytes, exceeding the {MaxAllocatedBytesPerRound:N0}-byte budget. This indicates an attacker-controlled size/count driving an allocation the configured limits never checked."));
            }
            if (stopwatch.Elapsed > MaxDurationPerRound)
            {
                throw new InvalidOperationException(
                    string.Create(CultureInfo.InvariantCulture,
                        $"Round took {stopwatch.Elapsed.TotalMilliseconds:N0}ms, exceeding the {MaxDurationPerRound.TotalMilliseconds:N0}ms budget. This indicates an unbounded loop or O(n^2) path reachable from untrusted input."));
            }
        }
    }

    // Wraps a byte array as a read-only, forward-only stream (CanSeek == false), for exercising the
    // non-seekable-source code paths (buffered .xls loading, CSV single-pass enumeration, etc.).
    internal sealed class NonSeekableStream : Stream
    {
        private readonly MemoryStream _inner;

        internal NonSeekableStream(byte[] bytes)
        {
            _inner = new MemoryStream(bytes);
        }

        public override bool CanRead => true;

        public override bool CanSeek => false;

        [ExcludeFromCodeCoverage]
        public override bool CanWrite => false;

        [ExcludeFromCodeCoverage]
        public override long Length => throw new NotSupportedException();

        [ExcludeFromCodeCoverage]
        public override long Position
        {
            get => _inner.Position; set => throw new NotSupportedException();
        }

        public override int Read(byte[] buffer, int offset, int count)
        {
            return _inner.Read(buffer, offset, count);
        }

        [ExcludeFromCodeCoverage]
        public override void Flush()
        {
        }

        [ExcludeFromCodeCoverage]
        public override long Seek(long offset, SeekOrigin origin)
        {
            throw new NotSupportedException();
        }

        [ExcludeFromCodeCoverage]
        public override void SetLength(long value)
        {
            throw new NotSupportedException();
        }

        [ExcludeFromCodeCoverage]
        public override void Write(byte[] buffer, int offset, int count)
        {
            throw new NotSupportedException();
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                _inner.Dispose();
            }
            base.Dispose(disposing);
        }
    }

    // Was duplicated (CoverageGapTests, XlsReaderTests) with a naming split (Disposed vs
    // WasDisposed) and a constructor split (parameterless vs byte[]). Both are kept here as overloads
    // rather than picking one, so neither call site needed to change its construction style.
    internal sealed class TrackingStream : MemoryStream
    {
        internal TrackingStream()
        {
        }

        internal TrackingStream(byte[] bytes)
            : base(bytes)
        {
        }

        internal bool Disposed { get; private set; }

        protected override void Dispose(bool disposing)
        {
            Disposed = true;
            base.Dispose(disposing);
        }
    }

    // Was duplicated identically in BufferedStreamCursorTests and MemoryZipParityTests
    // (GetSpan-backed, no Memory override — exercises the GetSpan-based read path). A third,
    // deliberately different copy in MemorySourceParityTests (Memory-backed, GetSpan throws — exercises
    // the opposite path on purpose) is NOT folded in here; merging it would silently change which code
    // path that test covers.
    internal sealed class NonArrayMemoryManager : MemoryManager<byte>
    {
        private readonly byte[] _data;

        internal NonArrayMemoryManager(byte[] data)
        {
            _data = data;
        }

        public override Span<byte> GetSpan()
        {
            return _data;
        }

        public override MemoryHandle Pin(int elementIndex = 0)
        {
            throw new NotSupportedException();
        }

        public override void Unpin()
        {
            throw new NotSupportedException();
        }

        [SuppressMessage("IDisposableAnalyzers.Correctness", "IDISP010:Call base.Dispose(disposing)",
            Justification = "MemoryManager<byte>.Dispose(bool) is abstract — there is no base implementation to call.")]
        protected override void Dispose(bool disposing)
        {
        }
    }

    // Was duplicated identically in MemoryZipParityTests and PrefetchDecompressionTests. A
    // third, deliberately different copy in SyncAsyncParityTests (adds StyleIndex, and stores
    // ValueBase64 instead of Value) is NOT folded in here for the same reason as NonArrayMemoryManager
    // above — it snapshots more state than these two need.
    internal readonly record struct CellSnapshot(
        int Row,
        int Column,
        int ColumnCount,
        CellType Type,
        string Value,
        bool HasDouble,
        long DoubleBits)
    {
        internal static CellSnapshot RowMarker(int row, int columnCount)
        {
            return new CellSnapshot(row, -1, columnCount, CellType.Empty, string.Empty, false, 0);
        }
    }

    // Marks a skipped (empty) column gap inside a TypedWorkbook row.
    // A record class (not struct) so `new Gap()` honors the Count = 1 default.
    internal sealed record Gap(int Count = 1);

    // Builds workbooks via the real XlsxWorkbookWriter from typed cell values.
    // Use for reader/parser fixtures expressible as inline strings, numbers,
    // dates (builtin numFmt 14), and bools. For shared strings, custom number
    // formats, the 1904 date system, or error/formula cells, use WorkbookBuilder
    // (raw XML) instead — XlsxWorkbookWriter cannot emit those.
    internal static class TypedWorkbook
    {
        // Single sheet "S1"; each row is an array of cell values.
        internal static Task<MemoryStream> BuildAsync(params object?[][] rows)
        {
            return BuildMultiSheetAsync(("S1", rows));
        }

        internal static async Task<MemoryStream> BuildMultiSheetAsync(
            params (string Name, object?[][] Rows)[] sheets)
        {
            var ms = new MemoryStream();
            await using (XlsxWorkbookWriter wb = await XlsxWorkbookWriter.CreateAsync(ms, leaveOpen: true))
            {
                await wb.StartAsync();
                foreach ((string name, object?[][] rows) in sheets)
                {
                    XlsxSheetWriter sheet = wb.AddSheet(name);
                    await sheet.StartAsync();
                    foreach (object?[] row in rows)
                    {
                        await using XlsxRowWriter rw = await sheet.StartRowAsync();
                        foreach (object? cell in row)
                        {
                            WriteCell(rw, cell);
                        }
                    }
                    await sheet.EndAsync();
                }
                await wb.EndAsync();
            }
            ms.Position = 0;
            return ms;
        }

        private static void WriteCell(XlsxRowWriter rw, object? cell)
        {
            switch (cell)
            {
                case null: rw.Write((string?)null); break;
                case string s: rw.Write(s); break;
                case bool b: rw.Write(b); break;
                case int i: rw.Write(i); break;
                case long l: rw.Write(l); break;
                case double d: rw.Write(d); break;
                case decimal m: rw.Write(m); break;
                case DateTime dt: rw.Write(dt); break;
                case Gap g: rw.Skip(g.Count); break;
                default: throw new NotSupportedException($"Unsupported cell value type: {cell.GetType()}");
            }
        }
    }

    internal static class WorkbookBuilder
    {
        private const string Main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        private const string Rel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        private const string PkgRel = "http://schemas.openxmlformats.org/package/2006/relationships";

        internal static MemoryStream Build(string sheetRows, string? sharedStrings = null, string? styles = null, bool date1904 = false)
        {
            return BuildMultiSheet([("S1", sheetRows)], sharedStrings, styles, date1904);
        }

        internal static MemoryStream BuildMultiSheet(
            (string Name, string Rows)[] sheets,
            string? sharedStrings = null,
            string? styles = null,
            bool date1904 = false)
        {
            var sheetXml = new string[sheets.Length];
            var relXml = new string[sheets.Length];
            var ms = new MemoryStream();
            using (var zip = new ZipArchive(ms, ZipArchiveMode.Create, leaveOpen: true))
            {
                for (int i = 0; i < sheets.Length; i++)
                {
                    int id = i + 1;
                    var (name, rows) = sheets[i];
                    sheetXml[i] = $"""<sheet name="{name}" sheetId="{id}" r:id="rId{id}"/>""";
                    relXml[i] = $"""<Relationship Id="rId{id}" Type="x" Target="worksheets/sheet{id}.xml"/>""";
                    Write(zip, $"xl/worksheets/sheet{id}.xml",
                        $"""<worksheet xmlns="{Main}"><sheetData>{rows}</sheetData></worksheet>""");
                }
                string workbookPr = date1904 ? """<workbookPr date1904="1"/>""" : "";
                Write(zip, "xl/workbook.xml",
                    $"""<workbook xmlns="{Main}" xmlns:r="{Rel}">{workbookPr}<sheets>{string.Concat(sheetXml)}</sheets></workbook>""");
                Write(zip, "xl/_rels/workbook.xml.rels",
                    $"""<Relationships xmlns="{PkgRel}">{string.Concat(relXml)}</Relationships>""");
                if (sharedStrings is not null)
                {
                    Write(zip, "xl/sharedStrings.xml", $"""<sst xmlns="{Main}">{sharedStrings}</sst>""");
                }
                if (styles is not null)
                {
                    string withNs = styles.Replace("<styleSheet>",
                        $"""<styleSheet xmlns="{Main}">""", StringComparison.Ordinal);
                    Write(zip, "xl/styles.xml", $"""<?xml version="1.0"?>{withNs}""");
                }
            }
            ms.Position = 0;
            return ms;
        }

        // Builds a workbook whose every SpreadsheetML element carries a namespace prefix (e.g. <x:row>),
        // as some non-Excel producers emit. The caller supplies already-prefixed row/shared/style content
        // this prefixes the structural elements (workbook/sheets/sheet/worksheet/sheetData/sst). The .rels
        // part keeps the OPC package-relationships namespace (never the spreadsheet prefix), matching reality.
        internal static MemoryStream BuildPrefixed(
            string prefix,
            string sheetRows,
            string? sharedStrings = null,
            string? stylesInner = null)
        {
            string p = prefix + ":";
            var ms = new MemoryStream();
            using (var zip = new ZipArchive(ms, ZipArchiveMode.Create, leaveOpen: true))
            {
                Write(zip, "xl/worksheets/sheet1.xml",
                    $"""<{p}worksheet xmlns:{prefix}="{Main}"><{p}sheetData>{sheetRows}</{p}sheetData></{p}worksheet>""");
                Write(zip, "xl/workbook.xml",
                    $"""<{p}workbook xmlns:{prefix}="{Main}" xmlns:r="{Rel}"><{p}sheets><{p}sheet name="S1" sheetId="1" r:id="rId1"/></{p}sheets></{p}workbook>""");
                Write(zip, "xl/_rels/workbook.xml.rels",
                    $"""<Relationships xmlns="{PkgRel}"><Relationship Id="rId1" Type="x" Target="worksheets/sheet1.xml"/></Relationships>""");
                if (sharedStrings is not null)
                {
                    Write(zip, "xl/sharedStrings.xml", $"""<{p}sst xmlns:{prefix}="{Main}">{sharedStrings}</{p}sst>""");
                }
                if (stylesInner is not null)
                {
                    Write(zip, "xl/styles.xml",
                        $"""<?xml version="1.0"?><{p}styleSheet xmlns:{prefix}="{Main}">{stylesInner}</{p}styleSheet>""");
                }
            }
            ms.Position = 0;
            return ms;
        }

        private static void Write(ZipArchive zip, string name, string content)
        {
            using var s = zip.CreateEntry(name).Open();
            byte[] bytes = Encoding.UTF8.GetBytes(content);
            s.Write(bytes, 0, bytes.Length);
        }
    }

    // An IUtf8SpanFormattable that formats as non-numeric text, so XlsbRowWriter.ToDouble's
    // final fallback (double.TryParse on the formatted bytes) fails and must throw rather than
    // silently return 0.0. Shared by XlsWriterTests and XlsbWriterTests since both formats route
    // Write<T> through the same XlsbRowWriter.ToDouble.
    internal readonly struct NonNumericFormattable : IUtf8SpanFormattable
    {
        public bool TryFormat(Span<byte> utf8Destination, out int bytesWritten, ReadOnlySpan<char> format, IFormatProvider? provider)
        {
            return Encoding.UTF8.TryGetBytes("not-a-number", utf8Destination, out bytesWritten);
        }
    }
}
