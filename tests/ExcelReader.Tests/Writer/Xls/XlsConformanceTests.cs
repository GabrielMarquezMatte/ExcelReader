using System.Buffers.Binary;
using ExcelReader.Core.Crypto;
using ExcelReader.Core.Reader;
using ExcelReader.Core.Writer.Xls;

namespace ExcelReader.Tests.Writer.Xls
{
    /// <summary>
    /// Excel refuses a BIFF8 file whose records break the [MS-XLS] grammar, while this library's own
    /// reader accepts it, so a round trip proves nothing here. These pin the structure Excel checks.
    /// </summary>
    public class XlsConformanceTests
    {
        private const int Bof = 0x0809;
        private const int Eof = 0x000A;
        private const int Index = 0x020B;
        private const int DefColWidth = 0x0055;
        private const int Row = 0x0208;
        private const int DbCell = 0x00D7;
        private const int Window1 = 0x003D;
        private const int Window2 = 0x023E;
        private const int WriteAccess = 0x005C;
        private const int Date1904 = 0x0022;
        private const int Label = 0x0204;
        private const int LabelSst = 0x00FD;
        private const int Sst = 0x00FC;
        private const int ExtSst = 0x00FF;

        private readonly record struct Record(int Position, int Id, byte[] Payload);

        private static byte[] Write(Action<XlsWorkbookWriter> build)
        {
            var ms = new MemoryStream();
            using (var wb = XlsWorkbookWriter.Create(ms, leaveOpen: true))
            {
                build(wb);
                wb.End();
            }
            return ms.ToArray();
        }

        private static List<Record> Records(byte[] xls)
        {
            using var cfb = CfbContainer.Parse(new MemoryStream(xls), ownsSource: true, ExcelReaderOptions.Default);
            byte[] stream = cfb.ReadStream("Workbook", maxBytes: 64 * 1024 * 1024);
            var records = new List<Record>();
            int pos = 0;
            while (pos + 4 <= stream.Length)
            {
                int id = BinaryPrimitives.ReadUInt16LittleEndian(stream.AsSpan(pos));
                if (id == 0)
                {
                    break;
                }
                int length = BinaryPrimitives.ReadUInt16LittleEndian(stream.AsSpan(pos + 2));
                records.Add(new Record(pos, id, stream.AsSpan(pos + 4, length).ToArray()));
                pos += 4 + length;
            }
            return records;
        }

        private static List<Record> Worksheet(List<Record> records)
        {
            int start = records.FindLastIndex(r => r.Id == Bof);
            return records.GetRange(start, records.Count - start);
        }

        [Fact]
        public void WorksheetSubstreamIndexesItsRowBlocks()
        {
            byte[] xls = Write(wb =>
            {
                using var sheet = wb.AddSheet("S");
                for (int i = 0; i < 40; i++)
                {
                    using var row = sheet.StartRow();
                    row.Write(i);
                    row.Write("text");
                }
            });

            List<Record> sheet = Worksheet(Records(xls));
            Dictionary<int, Record> byPosition = sheet.ToDictionary(r => r.Position);

            Assert.Equal(Bof, sheet[0].Id);
            Assert.Equal(Index, sheet[1].Id);
            Assert.Equal(Eof, sheet[^1].Id);
            Assert.Equal(Window2, sheet[^2].Id);

            byte[] index = sheet[1].Payload;
            Assert.Equal(40, BinaryPrimitives.ReadInt32LittleEndian(index.AsSpan(8)));
            Assert.Equal(DefColWidth, byPosition[BinaryPrimitives.ReadInt32LittleEndian(index.AsSpan(12))].Id);

            int blocks = (index.Length - 16) / 4;
            Assert.Equal(2, blocks);
            for (int block = 0; block < blocks; block++)
            {
                Record dbCell = byPosition[BinaryPrimitives.ReadInt32LittleEndian(index.AsSpan(16 + (4 * block)))];
                Assert.Equal(DbCell, dbCell.Id);
                int firstRow = dbCell.Position - BinaryPrimitives.ReadInt32LittleEndian(dbCell.Payload);
                Assert.Equal(Row, byPosition[firstRow].Id);
                Assert.Equal(block * 32, BinaryPrimitives.ReadUInt16LittleEndian(byPosition[firstRow].Payload));
            }

            Assert.Equal(40, sheet.Count(r => r.Id == Row));
            Assert.True(sheet.FindIndex(r => r.Id == Window2) > sheet.FindLastIndex(r => r.Id == DbCell));
        }

        [Fact]
        public void GlobalsRecordsHaveTheirSpecifiedShape()
        {
            byte[] xls = Write(wb =>
            {
                using var sheet = wb.AddSheet("S");
                using var row = sheet.StartRow();
                row.Write(1);
            });

            List<Record> records = Records(xls);
            Assert.Equal(0x07CC, BinaryPrimitives.ReadUInt16LittleEndian(records[0].Payload.AsSpan(6)));

            byte[] writeAccess = records.Single(r => r.Id == WriteAccess).Payload;
            Assert.Equal(112, writeAccess.Length);
            Assert.Equal("ExcelReader".Length, BinaryPrimitives.ReadUInt16LittleEndian(writeAccess));
            Assert.Equal(0, writeAccess[2]);

            Assert.Equal(18, records.Single(r => r.Id == Window1).Payload.Length);
            Assert.Equal([0, 0], records.Single(r => r.Id == Date1904).Payload);
        }

        [Fact]
        public void LongTextDifferingOnlyBetweenHashSamplesIsStillStoredOnce()
        {
            static string Text(int i)
            {
                return new string('a', 97) + i.ToString("D6", System.Globalization.CultureInfo.InvariantCulture) + new string('b', 197);
            }

            const int distinct = 200;
            byte[] xls = Write(wb =>
            {
                using var sheet = wb.AddSheet("S");
                for (int i = 0; i < 2 * distinct; i++)
                {
                    using var row = sheet.StartRow();
                    row.Write(Text(i % distinct));
                }
            });

            byte[] sst = Records(xls).Single(r => r.Id == Sst).Payload;
            Assert.Equal(2 * distinct, BinaryPrimitives.ReadInt32LittleEndian(sst));
            Assert.Equal(distinct, BinaryPrimitives.ReadInt32LittleEndian(sst.AsSpan(4)));

            using var reader = Excel.FromXls(new MemoryStream(xls));
            using var rows = reader.FirstSheet.GetEnumerator();
            for (int i = 0; i < 2 * distinct; i++)
            {
                Assert.True(rows.MoveNext());
                Assert.Equal(Text(i % distinct), rows.Current[0].GetString());
            }
            Assert.False(rows.MoveNext());
        }

        [Theory]
        [InlineData(255, 'x', Label)]
        [InlineData(256, 'x', LabelSst)]
        [InlineData(20_000, 'x', LabelSst)]
        [InlineData(9_000, 'é', LabelSst)]
        [InlineData(5_000, '日', LabelSst)]
        public void TextKeepsItsFullLength(int length, char fill, int expectedRecord)
        {
            string text = new(fill, length);
            byte[] xls = Write(wb =>
            {
                using var sheet = wb.AddSheet("S");
                using (var row = sheet.StartRow())
                {
                    row.Write(text);
                    row.Write(text);
                }
                using (var row = sheet.StartRow())
                {
                    row.Write("short");
                }
            });

            List<Record> records = Records(xls);
            Assert.Equal(2, Worksheet(records).Count(r => r.Id == expectedRecord));
            if (expectedRecord == LabelSst)
            {
                Assert.Single(records, r => r.Id == Sst);
                Assert.Single(records, r => r.Id == ExtSst);
                Assert.Equal([2, 0, 0, 0, 1, 0, 0, 0], records.Single(r => r.Id == Sst).Payload.AsSpan(0, 8).ToArray());
            }

            using var reader = Excel.FromXls(new MemoryStream(xls));
            using var rows = reader.FirstSheet.GetEnumerator();
            Assert.True(rows.MoveNext());
            Assert.Equal(text, rows.Current[0].GetString());
            Assert.Equal(text, rows.Current[1].GetString());
            Assert.True(rows.MoveNext());
            Assert.Equal("short", rows.Current[0].GetString());
            Assert.False(rows.MoveNext());
        }
    }
}
