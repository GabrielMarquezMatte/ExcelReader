using ExcelReader.Core.Reader;
using ExcelReader.Core.Reader.Xlsx;
using ExcelReader.Core.Writer.Xlsx;

namespace ExcelReader.Tests.Reader.Internal
{
    public sealed class SharedStringCacheTests
    {
        private static byte[] BuildXlsx()
        {
            using MemoryStream buffer = new();
            using (XlsxWorkbookWriter workbook = XlsxWorkbookWriter.Create(buffer, leaveOpen: true, new XlsxWriterOptions { UseSharedStrings = true }))
            {
                XlsxSheetWriter sheet = workbook.AddSheet("s");
                for (int r = 0; r < 50; r++)
                {
                    using XlsxRowWriter row = sheet.StartRow();
                    row.Write("value-" + (r % 5).ToString(System.Globalization.CultureInfo.InvariantCulture));
                }
                sheet.End();
                sheet.Dispose();
                workbook.End();
            }
            return buffer.ToArray();
        }

        private static List<string> FirstColumn(XlsxWorkbook reader)
        {
            List<string> values = [];
            using XlsxWorkbook.Enumerator e = reader.GetEnumerator();
            while (e.MoveNext())
            {
                values.Add(e.Current[0].GetString());
            }
            return values;
        }

        [Fact]
        public void A_Shared_String_Is_Materialized_Once_Per_Reader()
        {
            using XlsxWorkbook reader = Excel.FromXlsx(BuildXlsx());
            List<string> first = FirstColumn(reader);
            List<string> second = FirstColumn(reader);

            Assert.Equal(50, first.Count);
            for (int i = 0; i < first.Count; i++)
            {
                Assert.Same(first[i], second[i]);
                Assert.Same(first[i % 5], first[i]);
            }
        }

        [Fact]
        public void The_Cache_Array_Is_One_Instance()
        {
            using XlsxWorkbook reader = Excel.FromXlsx(BuildXlsx());
            _ = FirstColumn(reader);
            string?[] seen = reader.SharedStringCache;
            Parallel.For(0, 32, _ => Assert.Same(seen, reader.SharedStringCache));
        }
    }
}
