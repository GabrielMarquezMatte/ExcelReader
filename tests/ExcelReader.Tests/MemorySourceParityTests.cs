using System.Text;
using ExcelReader.Core.Reader;
using ExcelReader.Core.ValueObjects;

namespace ExcelReader.Tests
{
    public class MemorySourceParityTests
    {
        [Fact]
        public void CsvMemorySourceMatchesStreamSource()
        {
            byte[] bytes = Encoding.UTF8.GetBytes("name,age\nAda,37\n");

            using CsvReader streamReader = Excel.FromCsv(new MemoryStream(bytes, writable: false));
            using CsvReader memoryReader = Excel.FromCsv(bytes.AsMemory());

            string[] streamValues = ReadRows(streamReader);
            string[] memoryValues = ReadRows(memoryReader);

            Assert.Equal(streamValues, memoryValues);
        }

        [Fact]
        public void XlsMemorySourceMatchesStreamSource()
        {
            byte[] bytes = XlsWorkbookBuilder.Build(sheets: [("S1", [["Name", 1, true]])]).ToArray();

            using XlsReader streamReader = Excel.FromXls(new MemoryStream(bytes, writable: false));
            using XlsReader memoryReader = Excel.FromXls(bytes.AsMemory());

            string[] streamValues = ReadRows(streamReader);
            string[] memoryValues = ReadRows(memoryReader);

            Assert.Equal(streamValues, memoryValues);
        }

        private static string[] ReadRows(IExcelRowReader reader)
        {
            List<string> values = [];
            foreach (Row row in reader)
            {
                StringBuilder sb = new();
                foreach(var cell in row.Cells)
                {
                    sb.Append(cell.Value.GetString());
                    sb.Append('|');
                }
                values.Add(sb.ToString());
            }
            return [.. values];
        }
    }
}
