using System.Globalization;
using System.Runtime.InteropServices;
using System.Text;
using ExcelReader.Native;
using ExcelReader.Native.Typed;

namespace ExcelReader.Tests.Native
{
    public sealed class TypedParseSessionTests
    {
        private static readonly string XlsxFixture = Path.Combine(AppContext.BaseDirectory, "data", "sample.xlsx");

        private static NativeColumnSpec[] Specs()
        {
            return [new() { Index = 0, Type = NativeColumnType.String, Nullable = true }];
        }

        private static List<string> ReadInBatches(long maxRows, out int batchCount)
        {
            Assert.Equal(NativeStatus.Ok, NativeApiTests.OpenPath(XlsxFixture, NativeFormat.Auto, out NativeHandle? handle));
            using NativeHandle live = handle!;
            Assert.Equal(NativeStatus.Ok,
                TypedApi.TypedParseSession.Open(live, Specs(), headerRow: 0, maxRows, out TypedApi.TypedParseSession? session));
            using TypedApi.TypedParseSession open = session!;

            List<string> values = [];
            batchCount = 0;
            while (true)
            {
                int status = open.NextBatch(out NativeTable table);
                if (status == NativeStatus.Eof)
                {
                    Assert.Equal(0, table.ColumnCount);
                    Assert.Equal(IntPtr.Zero, table.Columns);
                    break;
                }
                Assert.Equal(NativeStatus.Ok, status);
                batchCount++;
                try
                {
                    values.AddRange(ReadStringColumn(table));
                }
                finally
                {
                    TypedApi.FreeTable(ref table);
                }
            }
            return values;
        }

        private static IEnumerable<string> ReadStringColumn(NativeTable table)
        {
            NativeColumn column = Marshal.PtrToStructure<NativeColumn>(table.Columns);
            for (long row = 0; row < column.Length; row++)
            {
                int start = Marshal.ReadInt32(column.Values, (int)(row * sizeof(int)));
                int end = Marshal.ReadInt32(column.Values, (int)((row + 1) * sizeof(int)));
                byte[] bytes = new byte[end - start];
                Marshal.Copy(IntPtr.Add(column.Data, start), bytes, 0, bytes.Length);
                yield return Encoding.UTF8.GetString(bytes);
            }
        }

        [Theory]
        [InlineData(1)]
        [InlineData(7)]
        [InlineData(8)]
        [InlineData(9)]
        [InlineData(1000)]
        public void NextBatch_Should_Match_The_Unbounded_Read_At_Every_Batch_Size(long maxRows)
        {
            List<string> whole = ReadInBatches(0, out int wholeBatches);
            List<string> batched = ReadInBatches(maxRows, out int batches);

            Assert.Equal(1, wholeBatches);
            Assert.Equal(whole, batched);
            long expectedBatches = (whole.Count + maxRows - 1) / maxRows;
            Assert.Equal(expectedBatches, batches);
        }

        private const int TallRowCount = 43;

        private readonly record struct TallRow(string Name, bool Active, bool QtyValid, long Qty);

        private static string WriteTallFixture(int rowCount = TallRowCount)
        {
            string path = Path.Combine(Path.GetTempPath(), $"excelreader-session-{Guid.NewGuid():N}.csv");
            StringBuilder csv = new("name,active,qty\n");
            for (int i = 0; i < rowCount; i++)
            {
                string active = i % 3 == 0 ? "true" : "false";
                string qty = i % 5 == 0 ? string.Empty : (i * 3).ToString(CultureInfo.InvariantCulture);
                csv.Append(CultureInfo.InvariantCulture, $"row-{i},{active},{qty}\n");
            }
            File.WriteAllText(path, csv.ToString());
            return path;
        }

        private static NativeColumnSpec[] TallSpecs()
        {
            return
            [
                new() { Names = ["name"], Type = NativeColumnType.String },
                new() { Names = ["active"], Type = NativeColumnType.Bool },
                new() { Names = ["qty"], Type = NativeColumnType.Int64, Nullable = true },
            ];
        }

        private static List<TallRow> ReadTallInBatches(string path, long maxRows, out int batchCount)
        {
            Assert.Equal(NativeStatus.Ok, NativeApiTests.OpenPath(path, NativeFormat.Csv, out NativeHandle? handle));
            using NativeHandle live = handle!;
            Assert.Equal(NativeStatus.Ok,
                TypedApi.TypedParseSession.Open(live, TallSpecs(), headerRow: 1, maxRows, out TypedApi.TypedParseSession? session));
            using TypedApi.TypedParseSession open = session!;

            List<TallRow> rows = [];
            batchCount = 0;
            while (true)
            {
                int status = open.NextBatch(out NativeTable table);
                if (status == NativeStatus.Eof)
                {
                    Assert.Equal(0, table.ColumnCount);
                    Assert.Equal(IntPtr.Zero, table.Columns);
                    break;
                }
                Assert.Equal(NativeStatus.Ok, status);
                batchCount++;
                try
                {
                    rows.AddRange(DecodeTallBatch(table));
                }
                finally
                {
                    TypedApi.FreeTable(ref table);
                }
            }
            return rows;
        }

        private static List<TallRow> DecodeTallBatch(NativeTable table)
        {
            NativeColumn names = ColumnAt(table, 0);
            NativeColumn active = ColumnAt(table, 1);
            NativeColumn qty = ColumnAt(table, 2);
            int rowCount = (int)names.Length;

            byte[] flags = new byte[rowCount];
            Marshal.Copy(active.Values, flags, 0, rowCount);
            long[] quantities = new long[rowCount];
            Marshal.Copy(qty.Values, quantities, 0, rowCount);
            bool[] valid = DecodeValidity(qty, rowCount);

            List<TallRow> decoded = [];
            for (int i = 0; i < rowCount; i++)
            {
                decoded.Add(new TallRow(ReadStringAt(names, i), flags[i] != 0, valid[i], quantities[i]));
            }
            return decoded;
        }

        private static NativeColumn ColumnAt(NativeTable table, int index)
        {
            return Marshal.PtrToStructure<NativeColumn>(IntPtr.Add(table.Columns, index * Marshal.SizeOf<NativeColumn>()));
        }

        private static string ReadStringAt(NativeColumn column, int row)
        {
            int start = Marshal.ReadInt32(column.Values, row * sizeof(int));
            int end = Marshal.ReadInt32(column.Values, (row + 1) * sizeof(int));
            byte[] bytes = new byte[end - start];
            Marshal.Copy(IntPtr.Add(column.Data, start), bytes, 0, bytes.Length);
            return Encoding.UTF8.GetString(bytes);
        }

        private static bool[] DecodeValidity(NativeColumn column, int rowCount)
        {
            bool[] result = new bool[rowCount];
            if (column.Validity == IntPtr.Zero)
            {
                Array.Fill(result, true);
                return result;
            }
            byte[] bitmap = new byte[(rowCount + 7) / 8];
            Marshal.Copy(column.Validity, bitmap, 0, bitmap.Length);
            for (int i = 0; i < rowCount; i++)
            {
                result[i] = (bitmap[i >> 3] & (1 << (i & 7))) != 0;
            }
            return result;
        }

        [Theory]
        [InlineData(1)]
        [InlineData(7)]
        [InlineData(8)]
        [InlineData(9)]
        [InlineData(1000)]
        public void NextBatch_Should_Match_The_Unbounded_Read_Across_Bit_Packing_Boundaries(long maxRows)
        {
            string path = WriteTallFixture();
            try
            {
                List<TallRow> whole = ReadTallInBatches(path, 0, out int wholeBatches);
                List<TallRow> batched = ReadTallInBatches(path, maxRows, out int batches);

                Assert.Equal(1, wholeBatches);
                Assert.Equal(TallRowCount, whole.Count);
                Assert.Contains(whole, row => row.Active);
                Assert.Contains(whole, row => !row.Active);
                Assert.Contains(whole, row => row.QtyValid);
                Assert.Contains(whole, row => !row.QtyValid);

                Assert.Equal(whole, batched);
                long expectedBatches = (whole.Count + maxRows - 1) / maxRows;
                Assert.Equal(expectedBatches, batches);
            }
            finally
            {
                File.Delete(path);
            }
        }

        private const int CeilingRowCount = 4000;
        private const long CeilingBatchSize = 200;

        private static long ColumnBytes(NativeColumn column)
        {
            long bytes = column.Type switch
            {
                NativeColumnType.String => (column.Length + 1) * sizeof(int) + column.DataLen,
                NativeColumnType.Bool => column.Length,
                NativeColumnType.Date => column.Length * sizeof(int),
                NativeColumnType.Float64 => column.Length * sizeof(double),
                _ => column.Length * sizeof(long),
            };
            if (column.Validity != IntPtr.Zero)
            {
                bytes += (column.Length + 7) / 8;
            }
            return bytes;
        }

        private static long TableBytes(NativeTable table)
        {
            long total = 0;
            for (int i = 0; i < table.ColumnCount; i++)
            {
                total += ColumnBytes(ColumnAt(table, i));
            }
            return total;
        }

        private static List<long> MeasureBatchByteSizes(string path, long maxRows)
        {
            Assert.Equal(NativeStatus.Ok, NativeApiTests.OpenPath(path, NativeFormat.Csv, out NativeHandle? handle));
            using NativeHandle live = handle!;
            Assert.Equal(NativeStatus.Ok,
                TypedApi.TypedParseSession.Open(live, TallSpecs(), headerRow: 1, maxRows, out TypedApi.TypedParseSession? session));
            using TypedApi.TypedParseSession open = session!;

            List<long> sizes = [];
            while (true)
            {
                int status = open.NextBatch(out NativeTable table);
                if (status == NativeStatus.Eof)
                {
                    break;
                }
                Assert.Equal(NativeStatus.Ok, status);
                try
                {
                    sizes.Add(TableBytes(table));
                }
                finally
                {
                    TypedApi.FreeTable(ref table);
                }
            }
            return sizes;
        }

        [Fact]
        public void NextBatch_Should_Bound_The_Largest_Live_Table_To_Roughly_One_Batch()
        {
            string path = WriteTallFixture(CeilingRowCount);
            try
            {
                List<long> unbounded = MeasureBatchByteSizes(path, maxRows: 0);
                List<long> batched = MeasureBatchByteSizes(path, CeilingBatchSize);

                Assert.Single(unbounded);
                Assert.True(batched.Count > 1, $"expected more than one batch, got {batched.Count}.");

                long unboundedBytes = unbounded[0];
                long maxBatchBytes = batched.Max();

                double expectedFraction = (double)CeilingBatchSize / CeilingRowCount;
                long bound = (long)(unboundedBytes * expectedFraction * 3);
                Assert.True(maxBatchBytes <= bound,
                    $"expected the largest live batch ({maxBatchBytes} bytes) to stay within {bound} " +
                    $"bytes (~{expectedFraction:P0} of the {unboundedBytes}-byte unbounded table, x3 " +
                    "slack) - the memory ceiling is not holding.");
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void OpenTypedReader_Should_Reject_A_Negative_Batch_Size()
        {
            Assert.Equal(NativeStatus.Ok, NativeApiTests.OpenPath(XlsxFixture, NativeFormat.Auto, out NativeHandle? handle));
            using NativeHandle live = handle!;

            int status = TypedApi.OpenTypedReader(live, Specs(), headerRow: 0, maxRows: -1, out nint reader);

            Assert.Equal(NativeStatus.InvalidArgument, status);
            Assert.Equal(0, reader);
        }

        [Fact]
        public void OpenTypedReader_Should_Reject_A_Null_Handle()
        {
            Assert.Equal(NativeStatus.InvalidHandle,
                TypedApi.OpenTypedReader(null, Specs(), headerRow: 0, maxRows: 10, out nint reader));
            Assert.Equal(0, reader);
        }

        [Fact]
        public void NextTypedBatch_Should_Report_InvalidHandle_For_A_Closed_Reader()
        {
            Assert.Equal(NativeStatus.Ok, NativeApiTests.OpenPath(XlsxFixture, NativeFormat.Auto, out NativeHandle? handle));
            using NativeHandle live = handle!;
            Assert.Equal(NativeStatus.Ok, TypedApi.OpenTypedReader(live, Specs(), 0, 10, out nint reader));

            TypedApi.CloseTypedReader(reader);

            Assert.Equal(NativeStatus.InvalidHandle, TypedApi.NextTypedBatch(reader, out NativeTable table));
            Assert.Equal(IntPtr.Zero, table.Columns);
        }

        [Fact]
        public void CloseTypedReader_Should_Be_A_No_Op_On_Zero_And_On_A_Second_Close()
        {
            Assert.Equal(NativeStatus.Ok, NativeApiTests.OpenPath(XlsxFixture, NativeFormat.Auto, out NativeHandle? handle));
            using NativeHandle live = handle!;
            Assert.Equal(NativeStatus.Ok, TypedApi.OpenTypedReader(live, Specs(), 0, 10, out nint reader));

            TypedApi.CloseTypedReader(reader);
            TypedApi.CloseTypedReader(reader);
            TypedApi.CloseTypedReader(0);
        }

        [Fact]
        public void NextTypedBatch_Should_Reject_A_Workbook_Id()
        {
            Assert.Equal(NativeStatus.Ok, NativeApiTests.OpenPath(XlsxFixture, NativeFormat.Auto, out NativeHandle? handle));
            using NativeHandle live = handle!;
            nint workbookId = NativeHandleTable.Register(live);

            Assert.Equal(NativeStatus.InvalidHandle, TypedApi.NextTypedBatch(workbookId, out _));
        }
    }
}
