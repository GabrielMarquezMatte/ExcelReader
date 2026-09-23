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

    internal sealed class TrackingStream : MemoryStream
    {
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

        protected override void Dispose(bool disposing)
        {
        }
    }

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

    internal sealed record Gap(int Count = 1);

    internal static class TypedWorkbook
    {
        internal static Task<MemoryStream> BuildAsync(params object?[][] rows)
        {
            return BuildMultiSheetAsync(("S1", rows));
        }

        internal static async Task<MemoryStream> BuildMultiSheetAsync(
            params (string Name, object?[][] Rows)[] sheets)
        {
            var ms = new MemoryStream();
            await using (XlsxWorkbookWriter wb = XlsxWorkbookWriter.Create(ms, leaveOpen: true))
            {
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

    internal readonly struct NonNumericFormattable : IUtf8SpanFormattable
    {
        public bool TryFormat(Span<byte> utf8Destination, out int bytesWritten, ReadOnlySpan<char> format, IFormatProvider? provider)
        {
            return Encoding.UTF8.TryGetBytes("not-a-number", utf8Destination, out bytesWritten);
        }
    }

    internal readonly struct OverflowingFormattable : IUtf8SpanFormattable
    {
        public bool TryFormat(Span<byte> utf8Destination, out int bytesWritten, ReadOnlySpan<char> format, IFormatProvider? provider)
        {
            return Encoding.UTF8.TryGetBytes("1e400", utf8Destination, out bytesWritten);
        }
    }
}
