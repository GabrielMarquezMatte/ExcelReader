using System.IO.Compression;
using ExcelReader.Core.Reader;
using ExcelReader.Core.Writer;

namespace ExcelReader.Tests
{
    // BrtColInfo is the one record this library writes that its own reader never parses — only
    // per-cell style matters for round-tripping — so no read-back test can catch a wrong field
    // layout here. These assert the emitted bytes directly.
    //
    // The expected layout was established by dumping a real Excel-authored .xlsb: an 18-byte payload
    // of colFirst/colLast/coldx/ixfe as u32 followed by flags as u16. This previously wrote a
    // 16-byte record (ixfe as u16), two bytes short of what Excel emits, with fUserSet left clear so
    // an explicit SetColumnWidth had no effect in Excel.
    public class XlsbColInfoTests
    {
        private const int ColInfoPayloadLength = 18;
        private const int FUserSet = 0x0002;

        private static async Task<byte[]> WriteSheetBinAsync(Action<XlsbSheetWriter> configure)
        {
            MemoryStream ms = new();
            await using (var wb = await XlsbWorkbookWriter.CreateAsync(ms, leaveOpen: true, ct: TestContext.Current.CancellationToken))
            {
                await wb.StartAsync(TestContext.Current.CancellationToken);
                XlsbSheetWriter sheet = wb.AddSheet("Sheet1");
                configure(sheet);
                await sheet.StartAsync(TestContext.Current.CancellationToken);
                await using (XlsbRowWriter row = await sheet.StartRowAsync(TestContext.Current.CancellationToken))
                {
                    row.Write("value");
                }
                await sheet.EndAsync(TestContext.Current.CancellationToken);
                await wb.EndAsync(TestContext.Current.CancellationToken);
            }
            ms.Position = 0;
            using ZipArchive zip = new(ms, ZipArchiveMode.Read);
            ZipArchiveEntry entry = Assert.Single(zip.Entries, e => string.Equals(e.FullName, "xl/worksheets/sheet1.bin", StringComparison.Ordinal));
            using MemoryStream inflated = new();
            using (Stream s = entry.Open())
            {
                await s.CopyToAsync(inflated, TestContext.Current.CancellationToken);
            }
            return inflated.ToArray();
        }

        private static List<byte[]> ReadColInfoPayloads(ReadOnlySpan<byte> data)
        {
            var payloads = new List<byte[]>();
            var reader = new Biff12RecordReader(data);
            while (reader.TryReadRecord(out int id, out ReadOnlySpan<byte> payload))
            {
                if (id == Brt.ColInfo)
                {
                    payloads.Add(payload.ToArray());
                }
            }
            return payloads;
        }

        [Fact]
        public async Task ColumnWidthEmitsEighteenByteRecordMatchingExcelLayout()
        {
            // 13.28515625 * 256 == 3401 exactly — the same column/width pair a real Excel file was
            // dumped from, so coldx is comparable byte-for-byte rather than rounding-dependent.
            byte[] sheet = await WriteSheetBinAsync(s => s.SetColumnWidth(4, 13.28515625));
            byte[] payload = Assert.Single(ReadColInfoPayloads(sheet));

            Assert.Equal(ColInfoPayloadLength, payload.Length);
            Assert.Equal(4u, BitConverter.ToUInt32(payload, 0));      // colFirst
            Assert.Equal(4u, BitConverter.ToUInt32(payload, 4));      // colLast
            Assert.Equal(3401u, BitConverter.ToUInt32(payload, 8));   // coldx
            Assert.Equal(0u, BitConverter.ToUInt32(payload, 12));     // ixfe (no style set)
            Assert.Equal(FUserSet, BitConverter.ToUInt16(payload, 16));
        }

        [Fact]
        public async Task ColumnStyleIsWrittenAsFourByteIxfe()
        {
            int styleId = 0;
            byte[] sheet = await WriteSheetBinAsync(s => s.SetColumnStyle(3, styleId));
            byte[] payload = Assert.Single(ReadColInfoPayloads(sheet));

            Assert.Equal(ColInfoPayloadLength, payload.Length);
            Assert.Equal(3u, BitConverter.ToUInt32(payload, 0));
            Assert.Equal(3u, BitConverter.ToUInt32(payload, 4));
            Assert.Equal((uint)styleId, BitConverter.ToUInt32(payload, 12));
        }

        [Fact]
        public async Task StyleWithoutExplicitWidthLeavesUserSetClear()
        {
            // Nothing claimed the width, so fUserSet must stay clear — otherwise Excel would treat the
            // library's 8.43 default as a deliberate user override of the sheet's own default width.
            byte[] sheet = await WriteSheetBinAsync(s => s.SetColumnStyle(2, 0));
            byte[] payload = Assert.Single(ReadColInfoPayloads(sheet));

            Assert.Equal(0, BitConverter.ToUInt16(payload, 16));
        }

        [Fact]
        public async Task EveryConfiguredColumnEmitsItsOwnRecordInAscendingOrder()
        {
            byte[] sheet = await WriteSheetBinAsync(s =>
            {
                s.SetColumnWidth(5, 20);
                s.SetColumnWidth(1, 10);
                s.SetColumnStyle(3, 0);
            });
            List<byte[]> payloads = ReadColInfoPayloads(sheet);

            Assert.Equal(3, payloads.Count);
            Assert.All(payloads, p => Assert.Equal(ColInfoPayloadLength, p.Length));
            Assert.Equal([1u, 3u, 5u], [.. payloads.Select(p => BitConverter.ToUInt32(p, 0))]);
            // Only the two columns given an explicit width claim fUserSet.
            Assert.Equal([FUserSet, 0, FUserSet], [.. payloads.Select(p => (int)BitConverter.ToUInt16(p, 16))]);
        }
    }
}
