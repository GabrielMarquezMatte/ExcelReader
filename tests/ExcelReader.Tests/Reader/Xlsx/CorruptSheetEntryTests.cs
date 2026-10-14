using System.Globalization;
using ExcelReader.Core.Reader;
using ExcelReader.Core.Reader.Xlsx;

namespace ExcelReader.Tests.Reader.Xlsx
{
    public class CorruptSheetEntryTests
    {
        private static async Task<byte[]> BuildAsync(int rows)
        {
            object?[][] data = new object?[rows + 1][];
            data[0] = ["Name", "Qty", "Day"];
            for (int i = 0; i < rows; i++)
            {
                data[i + 1] = [string.Create(CultureInfo.InvariantCulture, $"item{i:D6}"), i, new DateTime(2024, 1, 1).AddDays(i % 400)];
            }
            using MemoryStream ms = await TypedWorkbook.BuildAsync(data);
            return ms.ToArray();
        }

        private static int CountRows(XlsxWorkbook workbook)
        {
            int count = 0;
            foreach (Row _ in workbook.FirstSheet)
            {
                count++;
            }
            return count;
        }

        // The damage either breaks the deflate stream or decodes as garbage XML that trips a limit.
        // Either way the read must fail, not stop early without an error.
        [Theory]
        [InlineData(4096, false)]
        [InlineData(4096, true)]
        [InlineData(300_000, false)]
        public async Task ADamagedSheetEntryThrowsInsteadOfEndingEarly(int offset, bool prefetch)
        {
            byte[] bytes = await BuildAsync(30_000);
            int entry = bytes.AsSpan().IndexOf("xl/worksheets/sheet1.xml"u8);
            Array.Fill(bytes, (byte)0xFF, entry + offset, 512);
            using XlsxWorkbook workbook = Excel.FromXlsx(bytes, new ExcelReaderOptions { PrefetchDecompression = prefetch });

            Exception failure = Assert.ThrowsAny<Exception>(() => CountRows(workbook));
            Assert.True(failure is InvalidDataException or ExcelLimitExceededException, failure.GetType().Name);
        }
    }
}
