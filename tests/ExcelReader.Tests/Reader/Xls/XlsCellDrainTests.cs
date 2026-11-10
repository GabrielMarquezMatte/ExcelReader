using ExcelReader.Core.Reader;
using ExcelReader.Core.Reader.Internal;
using static ExcelReader.Tests.Reader.Xls.XlsWorkbookBuilder;

namespace ExcelReader.Tests.Reader.Xls
{
    public sealed class XlsCellDrainTests
    {
        private const int Number = 0x0203;
        private const int Rk = 0x027E;
        private const int MulRk = 0x00BD;

        private static byte[] RkInt(int row, int col, int value)
        {
            return [.. LE16(row), .. LE16(col), .. LE16(0), .. LE32((value << 2) | 2)];
        }

        private static byte[] MulRkInts(int row, int firstCol, params int[] values)
        {
            List<byte> payload = [.. LE16(row), .. LE16(firstCol)];
            foreach (int value in values)
            {
                payload.AddRange(LE16(0));
                payload.AddRange(LE32((value << 2) | 2));
            }
            payload.AddRange(LE16(firstCol + values.Length - 1));
            return [.. payload];
        }

        private static double[] ReadRow(Row row)
        {
            double[] values = new double[row.ColumnCount];
            for (int c = 0; c < values.Length; c++)
            {
                Assert.True(row[c].TryGetDouble(out values[c]));
            }
            return values;
        }

        [Fact]
        public void OutOfOrderNumberRecordsAreSortedWithinTheirRow()
        {
            using var ms = BuildRawSheet(
                includeEof: true,
                (Number, RawNumber(0, 3, 1.5)),
                (Rk, RkInt(0, 0, 7)),
                (MulRk, MulRkInts(0, 1, 10, 20)),
                (Number, RawNumber(1, 0, 2)));
            using var reader = Excel.FromXls(ms);
            using var e = reader.FirstSheet.GetEnumerator();

            Assert.True(e.MoveNext());
            Assert.Equal([7, 10, 20, 1.5], ReadRow(e.Current));
            Assert.True(e.MoveNext());
            Assert.Equal([2], ReadRow(e.Current));
            Assert.False(e.MoveNext());
        }

        [Fact]
        public void EmptyMulRkAndTruncatedNumberAddNoCells()
        {
            using var ms = BuildRawSheet(
                includeEof: true,
                (Rk, RkInt(0, 0, 1)),
                (MulRk, [.. LE16(0), .. LE16(1), .. new byte[6]]),
                (Number, [.. LE16(0), .. LE16(2), .. LE16(0), .. new byte[6]]),
                (Rk, RkInt(0, 3, 4)));
            using var reader = Excel.FromXls(ms);
            using var e = reader.FirstSheet.GetEnumerator();

            Assert.True(e.MoveNext());
            Assert.Equal(4, e.Current.ColumnCount);
            Assert.True(e.Current[0].TryGetDouble(out double first));
            Assert.Equal(1, first);
            Assert.Equal(CellType.Empty, e.Current[1].Type);
            Assert.Equal(CellType.Empty, e.Current[2].Type);
            Assert.True(e.Current[3].TryGetDouble(out double last));
            Assert.Equal(4, last);
            Assert.False(e.MoveNext());
        }

        [Fact]
        public void MulRkPastTheColumnLimitThrows()
        {
            using var ms = BuildRawSheet(
                includeEof: true,
                (Rk, RkInt(0, 0, 1)),
                (MulRk, MulRkInts(0, ExcelLimits.MaxColumns - 1, 1, 2)));
            using var reader = Excel.FromXls(ms);
            using var e = reader.FirstSheet.GetEnumerator();

            Assert.Throws<ExcelLimitExceededException>(() => e.MoveNext());
        }
    }
}
