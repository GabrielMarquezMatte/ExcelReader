using System.Text;
using ExcelReader.Core.Reader;
using B = ExcelReader.Tests.Biff12Build;

namespace ExcelReader.Tests
{
    public class XlsbPartsTests
    {
        // --- Biff12 field decoders ---

        [Fact]
        public void WideStringDecodes()
        {
            byte[] data = B.WideString("abc");
            Assert.True(Biff12.TryReadWideString(data, 0, out var chars, out int consumed));
            Assert.Equal("abc", chars.ToString());
            Assert.Equal(4 + 6, consumed);
        }

        [Fact]
        public void NullWideStringIsEmptyAndConsumesFourBytes()
        {
            byte[] data = B.NullWideString();
            Assert.True(Biff12.TryReadWideString(data, 0, out var chars, out int consumed));
            Assert.True(chars.IsEmpty);
            Assert.Equal(4, consumed);
        }

        [Fact]
        public void TruncatedWideStringReturnsFalse()
        {
            byte[] data = [.. B.U32(5), 1, 2, 3, 4]; // declares 5 chars (10 bytes), only 4 present
            Assert.False(Biff12.TryReadWideString(data, 0, out _, out _));
        }

        [Theory]
        [InlineData(402u, 100.0)]        // (100 << 2) | fInt
        [InlineData(0x3FF80000u, 1.5)]   // double 1.5, high 30 bits, fInt clear
        public void RkDecodesIntAndDouble(uint rk, double expected)
        {
            Assert.Equal(expected, Biff12.Rk(rk));
        }

        [Fact]
        public void RkAppliesDivideByHundred()
        {
            uint rk = (12345u << 2) | 0x03; // fInt | fX100
            Assert.Equal(123.45, Biff12.Rk(rk));
        }

        // --- workbook.bin ---

        private static byte[] WorkbookBin(uint wbPropFlags)
        {
            return
            [
                .. B.Record(Brt_WbProp, [.. B.U32(wbPropFlags), .. B.U32(0), .. B.WideString("")]),
                .. B.Record(Brt_BundleSh, [.. B.U32(0), .. B.U32(0), .. B.WideString("rId1"), .. B.WideString("Plan1")]),
                .. B.Record(Brt_BundleSh, [.. B.U32(0), .. B.U32(0), .. B.WideString("rId2"), .. B.WideString("Plan2")]),
            ];
        }

        private const int Brt_WbProp = 153;
        private const int Brt_BundleSh = 156;

        [Fact]
        public void ParsesSheetsWithRelTargets()
        {
            byte[] workbook = WorkbookBin(0);
            byte[] rels = Encoding.UTF8.GetBytes(
                """<Relationships><Relationship Id="rId1" Target="worksheets/sheet1.bin"/><Relationship Id="rId2" Target="/xl/worksheets/sheet2.bin"/></Relationships>""");

            var sheets = XlsbWorkbook.ParseSheets(workbook, rels);

            Assert.Equal(2, sheets.Length);
            Assert.Equal(("Plan1", "xl/worksheets/sheet1.bin"), sheets[0]);
            Assert.Equal(("Plan2", "xl/worksheets/sheet2.bin"), sheets[1]);
        }

        [Fact]
        public void ParsesDate1904Flag()
        {
            Assert.True(XlsbWorkbook.ParseDate1904(WorkbookBin(0x01)));
            Assert.False(XlsbWorkbook.ParseDate1904(WorkbookBin(0x00)));
            Assert.False(XlsbWorkbook.ParseDate1904([])); // no BrtWbProp
        }

        // --- styles.bin ---

        [Fact]
        public void StyleDateFlagsCoverBuiltinCustomAndCellXfsScope()
        {
            const int beginCellStyleXfs = 626;
            const int endCellStyleXfs = 627;
            byte[] styles =
            [
                .. B.Record(Brt.Fmt, [.. B.U16(176), .. B.WideString("yyyy-mm-dd")]), // custom date
                .. B.Record(Brt.Fmt, [.. B.U16(177), .. B.WideString("0.00")]),       // custom non-date
                // cellStyleXfs region: a date XF here must NOT count toward cell styles.
                .. B.Record(beginCellStyleXfs),
                .. B.Record(Brt.Xf, B.Xf(14)),
                .. B.Record(endCellStyleXfs),
                // cellXfs region: these are the ones iStyleRef indexes.
                .. B.Record(Brt.BeginCellXFs),
                .. B.Record(Brt.Xf, B.Xf(0)),    // general -> not date
                .. B.Record(Brt.Xf, B.Xf(14)),   // builtin date
                .. B.Record(Brt.Xf, B.Xf(176)),  // custom date
                .. B.Record(Brt.Xf, B.Xf(177)),  // custom non-date
                .. B.Record(Brt.EndCellXFs),
            ];

            bool[] flags = XlsbStyles.ParseStyleDateFlags(styles);

            Assert.Equal([false, true, true, false], flags);
        }

        // --- sharedStrings.bin ---

        [Fact]
        public void SharedStringsDecodeToFlatUtf8()
        {
            byte[] shared =
            [
                .. B.Record(Brt.SSTItem, [0, .. B.WideString("Alice")]),
                .. B.Record(Brt.SSTItem, [0, .. B.WideString("Café")]),
                .. B.Record(Brt.SSTItem, [0, .. B.WideString("Ω")]),
            ];

            var (flat, offsets) = XlsbSharedStrings.Parse(shared);

            Assert.Equal(4, offsets.Length);
            Assert.Equal("Alice", At(flat, offsets, 0));
            Assert.Equal("Café", At(flat, offsets, 1));
            Assert.Equal("Ω", At(flat, offsets, 2));
        }

        private static string At(byte[] flat, int[] offsets, int index)
        {
            return Encoding.UTF8.GetString(flat.AsSpan(offsets[index], offsets[index + 1] - offsets[index]));
        }
    }
}
