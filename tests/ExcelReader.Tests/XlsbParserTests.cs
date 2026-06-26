using System.IO.Compression;
using System.Text;
using ExcelReader.Core.Parser;
using ExcelReader.Core.Reader;
using B = ExcelReader.Tests.Biff12Build;

namespace ExcelReader.Tests
{
    public class XlsbParserTests
    {
        private sealed class PersonRow
        {
            public string? Name { get; set; }
            public int Age { get; set; }
            public bool Active { get; set; }
        }

        [Fact]
        public void ParserMapsXlsbRows()
        {
            using var ms = BuildWorkbook();
            using var reader = Excel.FromXlsb(ms);

            var rows = new ExcelParser<PersonRow>().Parse(reader).ToList();

            Assert.Single(rows);
            Assert.Equal("Alice", rows[0].Name);
            Assert.Equal(42, rows[0].Age);
            Assert.True(rows[0].Active);
        }

        [Fact]
        public async Task AsyncParserMapsXlsbRows()
        {
            await using var ms = BuildWorkbook();
            await using var reader = await Excel.FromXlsbAsync(ms, ct: TestContext.Current.CancellationToken);
            var rows = new List<PersonRow>();

            await foreach (var row in new ExcelParser<PersonRow>().ParseAsync(reader, TestContext.Current.CancellationToken))
            {
                rows.Add(row);
            }

            Assert.Single(rows);
            Assert.Equal("Alice", rows[0].Name);
            Assert.Equal(42, rows[0].Age);
            Assert.True(rows[0].Active);
        }

        private static MemoryStream BuildWorkbook()
        {
            MemoryStream ms = new();
            using (var zip = new ZipArchive(ms, ZipArchiveMode.Create, leaveOpen: true))
            {
                Add(zip, "xl/workbook.bin",
                [
                    .. B.Record(Brt.WbProp, [.. B.U32(0), .. B.U32(0), .. B.WideString("")]),
                    .. B.Record(Brt.BundleSh, [.. B.U32(0), .. B.U32(0), .. B.WideString("rId1"), .. B.WideString("Sheet1")]),
                ]);
                Add(zip, "xl/_rels/workbook.bin.rels", Encoding.UTF8.GetBytes(
                    """<Relationships><Relationship Id="rId1" Target="worksheets/sheet1.bin"/></Relationships>"""));
                Add(zip, "xl/styles.bin", []);
                Add(zip, "xl/sharedStrings.bin", []);
                Add(zip, "xl/worksheets/sheet1.bin",
                [
                    .. B.Record(Brt.RowHdr),
                    .. B.Record(Brt.CellSt, B.CellSt(0, 0, "Name")),
                    .. B.Record(Brt.CellSt, B.CellSt(1, 0, "Age")),
                    .. B.Record(Brt.CellSt, B.CellSt(2, 0, "Active")),
                    .. B.Record(Brt.RowHdr),
                    .. B.Record(Brt.CellSt, B.CellSt(0, 0, "Alice")),
                    .. B.Record(Brt.CellRk, B.CellRk(1, 0, (42u << 2) | 0x02)),
                    .. B.Record(Brt.CellBool, B.CellBool(2, 0, true)),
                    .. B.Record(Brt.EndSheetData),
                ]);
            }
            ms.Position = 0;
            return ms;
        }

        private static void Add(ZipArchive zip, string name, byte[] bytes)
        {
            var entry = zip.CreateEntry(name);
            using var stream = entry.Open();
            stream.Write(bytes);
        }
    }
}