using System.Globalization;
using ExcelReader.Core.Writer.Xlsb;
using ExcelReader.Native;
using ExcelReader.Native.Reading;
using ExcelReader.Native.Typed;

namespace ExcelReader.Tests.Native
{
    public sealed partial class NativeApiTests
    {
        private static object?[][] MixedSheetRows(int rows)
        {
            object?[][] data = new object?[rows + 1][];
            data[0] = ["name", "qty", "price", "flag", "day", "note"];
            for (int i = 0; i < rows; i++)
            {
                data[i + 1] =
                [
                    string.Create(CultureInfo.InvariantCulture, $"item{i}"),
                    i % 7 == 0 ? null : i * 3,
                    i * 0.25,
                    i % 2 == 0,
                    new DateTime(2024, 1, 1).AddDays(i % 28),
                    string.Create(CultureInfo.InvariantCulture, $"n{i}"),
                ];
            }
            return data;
        }

        private static async Task<byte[]> MixedXlsbAsync(int rows)
        {
            CancellationToken ct = TestContext.Current.CancellationToken;
            var ms = new MemoryStream();
            await using (XlsbWorkbookWriter wb = XlsbWorkbookWriter.Create(ms, leaveOpen: true))
            {
                XlsbSheetWriter sheet = wb.AddSheet("S1");
                foreach (object?[] cells in MixedSheetRows(rows))
                {
                    await using XlsbRowWriter row = await sheet.StartRowAsync(ct);
                    foreach (object? cell in cells)
                    {
                        switch (cell)
                        {
                            case null: row.Write((string?)null); break;
                            case string s: row.Write(s); break;
                            case bool b: row.Write(b); break;
                            case int i: row.Write(i); break;
                            case double d: row.Write(d); break;
                            case DateTime dt: row.Write(dt); break;
                        }
                    }
                }
                await sheet.EndAsync(ct);
                await wb.EndAsync(ct);
            }
            return ms.ToArray();
        }

        private static (int Status, NativeTable Table, bool Parallel) ParseSheetWith(byte[] workbook, int format, int dop, int chunkBytes)
        {
            Assert.Equal(NativeStatus.Ok, ReadApi.OpenMemory(workbook, format, out NativeHandle? handle));
            try
            {
                int status = TypedApi.ParseTypedTable(handle, 0, MixedSpecs, 1, dop, out NativeTable table, chunkBytes);
                return (status, table, TypedApi.LastParseRanInParallel);
            }
            finally
            {
                ReadApi.Close(handle);
            }
        }

        [Theory]
        [InlineData(2, 0)]
        [InlineData(8, 0)]
        [InlineData(4, 4096)]
        public async Task ParallelXlsxTypedParseMatchesSequential(int dop, int chunkBytes)
        {
            using MemoryStream built = await TypedWorkbook.BuildAsync(MixedSheetRows(20_000));
            byte[] workbook = built.ToArray();
            (int expectedStatus, NativeTable expected, bool expectedParallel) = ParseSheetWith(workbook, NativeFormat.Xlsx, 1, 0);
            (int status, NativeTable actual, bool parallel) = ParseSheetWith(workbook, NativeFormat.Xlsx, dop, chunkBytes);
            try
            {
                Assert.Equal(NativeStatus.Ok, expectedStatus);
                Assert.Equal(NativeStatus.Ok, status);
                Assert.False(expectedParallel);
                Assert.True(parallel);
                AssertSameTable(expected, actual);
            }
            finally
            {
                TypedApi.FreeTable(ref expected);
                TypedApi.FreeTable(ref actual);
            }
        }

        [Theory]
        [InlineData(2, 0)]
        [InlineData(8, 0)]
        [InlineData(4, 4096)]
        public async Task ParallelXlsbTypedParseMatchesSequential(int dop, int chunkBytes)
        {
            byte[] workbook = await MixedXlsbAsync(40_000);
            (int expectedStatus, NativeTable expected, _) = ParseSheetWith(workbook, NativeFormat.Xlsb, 1, 0);
            (int status, NativeTable actual, bool parallel) = ParseSheetWith(workbook, NativeFormat.Xlsb, dop, chunkBytes);
            try
            {
                Assert.Equal(NativeStatus.Ok, expectedStatus);
                Assert.Equal(NativeStatus.Ok, status);
                Assert.True(parallel);
                AssertSameTable(expected, actual);
            }
            finally
            {
                TypedApi.FreeTable(ref expected);
                TypedApi.FreeTable(ref actual);
            }
        }

        [Fact]
        public async Task AnInvalidSheetIndexStillFailsTheSameWay()
        {
            using MemoryStream built = await TypedWorkbook.BuildAsync(MixedSheetRows(20_000));
            Assert.Equal(NativeStatus.Ok, ReadApi.OpenMemory(built.ToArray(), NativeFormat.Xlsx, out NativeHandle? handle));
            try
            {
                int sequential = TypedApi.ParseTypedTable(handle, 9, MixedSpecs, 1, 1, out _);
                int parallel = TypedApi.ParseTypedTable(handle, 9, MixedSpecs, 1, 8, out _);

                Assert.NotEqual(NativeStatus.Ok, sequential);
                Assert.Equal(sequential, parallel);
            }
            finally
            {
                ReadApi.Close(handle);
            }
        }
    }
}
