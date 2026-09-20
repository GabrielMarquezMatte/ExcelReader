using System.Text;
using ExcelReader.Core.Parser;
using ExcelReader.Core.Reader;
using ExcelReader.Core.ValueObjects;

namespace ExcelReader.Tests
{
    public class ExcelCellReadersUtf8Tests
    {
        [Fact]
        public void ReadsTheRawBytesOfEveryField()
        {
            byte[] csv = "plain,\"quoted, \"\"text\"\"\",\n"u8.ToArray();
            using CsvReader reader = Excel.FromCsv(csv);
            var seen = new List<string>();
            foreach (Row row in reader)
            {
                for (int i = 0; i < 3; i++)
                {
                    Assert.True(ExcelCellReaders.Utf8(row[i], false, System.Globalization.CultureInfo.InvariantCulture, out ReadOnlySpan<byte> value));
                    seen.Add(Encoding.UTF8.GetString(value));
                }
            }
            Assert.Equal(["plain", "quoted, \"text\"", ""], seen);
        }
    }
}
