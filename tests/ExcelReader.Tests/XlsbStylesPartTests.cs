using System.IO.Compression;
using ExcelReader.Core.Reader;
using ExcelReader.Core.Writer;

namespace ExcelReader.Tests
{
    // The styles part is structural: this library's own reader only consumes the number-format side of
    // it, so a missing or malformed production round-trips fine here while Excel reports "Formato de
    // parte de /xl/styles.bin" and repairs the workbook on open. These assert the record structure
    // directly, against the shape a real Excel-authored .xlsb was dumped from.
    public class XlsbStylesPartTests
    {
        private static async Task<byte[]> WriteStylesBinAsync(Action<XlsbWorkbookWriter>? configure = null)
        {
            MemoryStream ms = new();
            await using (var wb = await XlsbWorkbookWriter.CreateAsync(ms, leaveOpen: true, ct: TestContext.Current.CancellationToken))
            {
                await wb.StartAsync(TestContext.Current.CancellationToken);
                configure?.Invoke(wb);
                XlsbSheetWriter sheet = wb.AddSheet("Sheet1");
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
            ZipArchiveEntry entry = Assert.Single(zip.Entries, e => string.Equals(e.FullName, "xl/styles.bin", StringComparison.Ordinal));
            using MemoryStream inflated = new();
            using (Stream s = entry.Open())
            {
                await s.CopyToAsync(inflated, TestContext.Current.CancellationToken);
            }
            return inflated.ToArray();
        }

        private static List<(int Id, byte[] Payload)> ReadRecords(ReadOnlySpan<byte> data)
        {
            var records = new List<(int, byte[])>();
            var reader = new Biff12RecordReader(data);
            while (reader.TryReadRecord(out int id, out ReadOnlySpan<byte> payload))
            {
                records.Add((id, payload.ToArray()));
            }
            return records;
        }

        [Fact]
        public async Task StylesPartContainsMandatoryStylesProduction()
        {
            List<(int Id, byte[] Payload)> records = ReadRecords(await WriteStylesBinAsync());
            List<int> ids = [.. records.Select(r => r.Id)];

            Assert.Contains(Brt.BeginStyles, ids);
            Assert.Contains(Brt.Style, ids);
            Assert.Contains(Brt.EndStyles, ids);
        }

        [Fact]
        public async Task StylesProductionSitsBetweenCellXfsAndEndOfStyleSheet()
        {
            List<int> ids = [.. ReadRecords(await WriteStylesBinAsync()).Select(r => r.Id)];

            int endCellXfs = ids.IndexOf(Brt.EndCellXFs);
            int beginStyles = ids.IndexOf(Brt.BeginStyles);
            int endStyles = ids.IndexOf(Brt.EndStyles);
            int endStyleSheet = ids.IndexOf(Brt.EndStyleSheet);

            Assert.True(endCellXfs >= 0 && beginStyles >= 0 && endStyles >= 0 && endStyleSheet >= 0);
            Assert.True(endCellXfs < beginStyles, "STYLES must follow CELLXFS.");
            Assert.True(beginStyles < endStyles, "BrtBeginStyles must precede BrtEndStyles.");
            Assert.True(endStyles < endStyleSheet, "STYLES must close before BrtEndStyleSheet.");
        }

        [Fact]
        public async Task NormalStyleRecordMatchesExcelByteLayout()
        {
            List<(int Id, byte[] Payload)> records = ReadRecords(await WriteStylesBinAsync());

            byte[] count = Assert.Single(records, r => r.Id == Brt.BeginStyles).Payload;
            Assert.Equal(1u, BitConverter.ToUInt32(count, 0));

            byte[] style = Assert.Single(records, r => r.Id == Brt.Style).Payload;
            // Byte-for-byte identical to the "Normal" BrtStyle in a real Excel-authored .xlsb:
            //   ixf=0 | grbit=0x0001 (fBuiltIn) | iStyBuiltIn=0 | iLevel=0 | "Normal" as XLWideString
            Assert.Equal(
                Convert.FromHexString("0000000001000000060000004E006F0072006D0061006C00"),
                style);
        }

        [Fact]
        public async Task FillsAreTheTwoExcelAlwaysRequires()
        {
            List<(int Id, byte[] Payload)> records = ReadRecords(await WriteStylesBinAsync());

            byte[] count = Assert.Single(records, r => r.Id == Brt.BeginFills).Payload;
            Assert.Equal(2u, BitConverter.ToUInt32(count, 0));

            List<byte[]> fills = [.. records.Where(r => r.Id == Brt.Fill).Select(r => r.Payload)];
            Assert.Equal(2, fills.Count);
            // Byte-for-byte identical to the two fills in a real Excel-authored .xlsb. The payload used
            // to be written one byte short at the head, shifting every field left so the leading fls
            // (pattern type) decoded as 0x03000000 instead of 0.
            Assert.Equal(
                Convert.FromHexString("0000000003400000000000FF03410000FFFFFFFF" + new string('0', 96)),
                fills[0]);
            Assert.Equal(
                Convert.FromHexString("1100000003400000000000FF03410000FFFFFFFF" + new string('0', 96)),
                fills[1]);
        }

        [Fact]
        public async Task XfRecordsAreFullyInitialized()
        {
            List<byte[]> xfs = [.. ReadRecords(await WriteStylesBinAsync()).Where(r => r.Id == Brt.Xf).Select(r => r.Payload)];

            // The middle 10 bytes (iFont/iFill/ixBorder/trot/indent/flags) used to come from an
            // uninitialized `stackalloc` under [SkipLocalsInit], so they held arbitrary stack contents
            // — indices pointing at fonts and fills the part never declares. Both of these are
            // byte-for-byte what Excel writes for the equivalent XF.
            Assert.Equal(
                Convert.FromHexString("FFFF0000000000000000000010100000"), // cellStyleXfs[0]
                xfs[0]);
            Assert.Equal(
                Convert.FromHexString("00000000000000000000000010100000"), // cellXfs[0], no number format
                xfs[1]);
            Assert.All(xfs, xf => Assert.Equal(16, xf.Length));
            // iFont/iFill/ixBorder must stay 0: the part declares exactly one font and one border.
            Assert.All(xfs, xf =>
            {
                Assert.Equal(0, BitConverter.ToUInt16(xf, 4));  // iFont
                Assert.Equal(0, BitConverter.ToUInt16(xf, 6));  // iFill
                Assert.Equal(0, BitConverter.ToUInt16(xf, 8));  // ixBorder
            });
        }

        [Fact]
        public async Task CustomNumberFormatXfFlagsTheOverriddenAttribute()
        {
            List<(int Id, byte[] Payload)> records = ReadRecords(
                await WriteStylesBinAsync(wb => wb.AddStyle(new CellStyle { NumberFormat = "0.00%" })));
            List<byte[]> xfs = [.. records.Where(r => r.Id == Brt.Xf).Select(r => r.Payload)];

            // The last cell XF is the custom style's: it carries a non-zero iFmt, and must therefore
            // set xfGrbitAtr's number-format bit, exactly as Excel does.
            byte[] custom = xfs[^1];
            Assert.NotEqual(0, BitConverter.ToUInt16(custom, 2));  // iFmt
            Assert.Equal(1, BitConverter.ToUInt16(custom, 14));    // xfGrbitAtr: fAtrNum
        }

        [Fact]
        public async Task StylesProductionSurvivesCustomStyles()
        {
            // A custom number format grows CELLXFS; the STYLES production must still be emitted after it.
            List<int> ids = [.. ReadRecords(await WriteStylesBinAsync(wb => wb.AddStyle(new CellStyle { NumberFormat = "0.00%" }))).Select(r => r.Id)];

            Assert.True(ids.IndexOf(Brt.EndCellXFs) < ids.IndexOf(Brt.BeginStyles));
            Assert.Contains(Brt.Style, ids);
            Assert.True(ids.IndexOf(Brt.EndStyles) < ids.IndexOf(Brt.EndStyleSheet));
        }
    }
}
