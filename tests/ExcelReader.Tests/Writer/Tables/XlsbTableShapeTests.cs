using ExcelReader.Core.Reader.Xlsb;
using ExcelReader.Core.Writer;
using B = ExcelReader.Tests.Reader.Xlsb.Biff12Build;

namespace ExcelReader.Tests.Writer.Tables
{
    public class XlsbTableShapeTests
    {
        [Fact]
        public void TablePartHasExcelsRecordSequence()
        {
            List<(int Id, byte[] Payload)> records = Records(TableBooks.Entry(StockPackage(), "xl/tables/table1.bin"));

            Assert.Equal([343, 161, 162, 345, 347, 348, 347, 348, 346, 513, 344], records.Select(static r => r.Id));
        }

        [Fact]
        public void BeginListMatchesExcelsLayout()
        {
            byte[] payload = Records(TableBooks.Entry(StockPackage(), "xl/tables/table1.bin"))[0].Payload;

            byte[] expected =
            [
                .. B.U32(0), .. B.U32(2), .. B.U32(0), .. B.U32(1),
                .. B.U32(0), .. B.U32(1), .. B.U32(1), .. B.U32(0), .. B.U32(0),
                .. B.NullWideString(), .. B.NullWideString(), .. B.NullWideString(),
                .. B.NullWideString(), .. B.NullWideString(), .. B.NullWideString(),
                .. B.U32(0),
                .. B.WideString("Stock"), .. B.WideString("Stock"), .. B.WideString(""),
                .. B.NullWideString(), .. B.NullWideString(), .. B.NullWideString(),
            ];
            Assert.Equal(108, payload.Length);
            Assert.Equal(expected, payload);
        }

        [Fact]
        public void ColumnsAutoFilterAndStyleMatchExcelsLayout()
        {
            List<(int Id, byte[] Payload)> records = Records(TableBooks.Entry(StockPackage(), "xl/tables/table1.bin"));

            Assert.Equal([.. B.U32(0), .. B.U32(2), .. B.U32(0), .. B.U32(1)], records[1].Payload);
            Assert.Equal(B.U32(2), records[3].Payload);
            Assert.Equal(
                [
                    .. B.U32(1), .. B.U32(0), .. B.NullWideString(), .. B.NullWideString(), .. B.NullWideString(), .. B.U32(0),
                    .. B.NullWideString(), .. B.WideString("Product"),
                    .. B.NullWideString(), .. B.NullWideString(), .. B.NullWideString(), .. B.NullWideString(),
                ],
                records[4].Payload);
            Assert.Equal([.. B.U16(4), .. B.WideString("TableStyleMedium2")], records[9].Payload);
        }

        [Fact]
        public void StyleFlagsAndNullStyleAreEncoded()
        {
            ITableBook book = TableBooks.Create("xlsb");
            book.AddSheet("S");
            book.BeginTable("Plain", ["H"], new ExcelTableOptions
            {
                StyleName = null,
                ShowFirstColumn = true,
                ShowLastColumn = true,
                ShowRowStripes = false,
                ShowColumnStripes = true,
            });

            byte[] style = Records(TableBooks.Entry(book.Finish(), "xl/tables/table1.bin"))[^2].Payload;

            Assert.Equal([.. B.U16(1 | 2 | 8), .. B.NullWideString()], style);
        }

        [Fact]
        public void SheetListsItsTablePartsBeforeEndSheet()
        {
            byte[] package = StockPackage();

            List<(int Id, byte[] Payload)> sheet = Records(TableBooks.Entry(package, "xl/worksheets/sheet1.bin"));

            Assert.Equal([660, 661, 662, 130], sheet.Skip(sheet.Count - 4).Select(static r => r.Id));
            Assert.Equal(B.U32(1), sheet[^4].Payload);
            Assert.Equal(B.WideString("rId1"), sheet[^3].Payload);
            string types = System.Text.Encoding.UTF8.GetString(TableBooks.Entry(package, "[Content_Types].xml"));
            Assert.Contains("<Override PartName=\"/xl/tables/table1.bin\" ContentType=\"application/vnd.ms-excel.table\"/>", types, StringComparison.Ordinal);
            Assert.Contains("Target=\"../tables/table1.bin\"", System.Text.Encoding.UTF8.GetString(TableBooks.Entry(package, "xl/worksheets/_rels/sheet1.bin.rels")), StringComparison.Ordinal);
        }

        [Fact]
        public void SheetWithoutTableHasNoAutoFilterOrSortState()
        {
            ITableBook book = TableBooks.Create("xlsb");
            book.AddSheet("S");
            book.Row("Pen", 100);

            List<int> ids = [.. Records(TableBooks.Entry(book.Finish(), "xl/worksheets/sheet1.bin")).Select(static r => r.Id)];

            Assert.DoesNotContain(161, ids);
            Assert.DoesNotContain(648, ids);
            Assert.Equal([146, 130], ids.Skip(ids.Count - 2));
        }

        [Fact]
        public void TableSheetTailAfterSheetDataIsOnlyListParts()
        {
            List<int> ids = [.. Records(TableBooks.Entry(StockPackage(), "xl/worksheets/sheet1.bin")).Select(static r => r.Id)];

            Assert.Equal([660, 661, 662, 130], ids.Skip(ids.IndexOf(146) + 1));
        }

        private static byte[] StockPackage()
        {
            ITableBook book = TableBooks.Create("xlsb");
            book.AddSheet("S");
            book.BeginTable("Stock", ["Product", "Qty"]);
            book.Row("Pen", 100);
            book.Row("Notebook", 40);
            return book.Finish();
        }

        private static List<(int Id, byte[] Payload)> Records(byte[] part)
        {
            List<(int, byte[])> records = [];
            Biff12RecordReader reader = new(part);
            while (reader.TryReadRecord(out int id, out ReadOnlySpan<byte> payload))
            {
                records.Add((id, payload.ToArray()));
            }
            return records;
        }
    }
}
