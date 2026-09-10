using System.Globalization;
using System.Runtime.InteropServices;
using System.Text;
using ExcelReader.Native;

namespace ExcelReader.Tests
{
    public sealed class TypedParseSessionTests
    {
        private static readonly string XlsxFixture = Path.Combine(AppContext.BaseDirectory, "data", "sample.xlsx");

        private static NativeColumnSpec[] Specs()
        {
            return [new() { Index = 0, Type = NativeColumnType.String, Nullable = true }];
        }

        // Reads every batch at `maxRows` and returns the flattened per-row strings of column 0.
        private static List<string> ReadInBatches(long maxRows, out int batchCount)
        {
            Assert.Equal(NativeStatus.Ok, NativeApiTests.OpenPath(XlsxFixture, NativeFormat.Auto, out NativeHandle? handle));
            using NativeHandle live = handle!;
            Assert.Equal(NativeStatus.Ok,
                NativeApi.TypedParseSession.Open(live, Specs(), headerRow: 0, maxRows, out NativeApi.TypedParseSession? session));
            using NativeApi.TypedParseSession open = session!;

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
                    NativeApi.FreeTable(ref table);
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

        // The load-bearing property: batching never changes the data. Sizes that are not multiples
        // of 8 are deliberate — the validity and bool buffers are LSB-first bit-packed, so a batch
        // ending mid-byte is exactly where a packing bug surfaces.
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
            // Concatenated values alone cannot tell batching apart from a NextBatch that ignored
            // maxRows and returned the whole sheet in one call, so the batch count is pinned too:
            // ceil(rows / maxRows), with the row count taken from the unbounded read so this stays
            // correct if the fixture grows.
            long expectedBatches = (whole.Count + maxRows - 1) / maxRows;
            Assert.Equal(expectedBatches, batches);
        }

        // sample.xlsx holds 3 rows, so on it every batch size above 2 collapses to a single batch and
        // no batch ever ends mid-byte. This second sweep runs the same property over a fixture tall
        // enough that 7, 8 and 9 each produce several batches plus a partial final one — 43 is
        // coprime with all three, so no size divides it evenly.
        private const int TallRowCount = 43;

        // One decoded row of the tall fixture. Compared as a whole so a bit-packing bug in the bool
        // values or the validity bitmap fails the assertion, which a string-only comparison would miss.
        private readonly record struct TallRow(string Name, bool Active, bool QtyValid, long Qty);

        // A temp CSV is exactly how NativeApiTests.cs builds its own ParseTyped fixtures
        // (ParseTyped_Should_Return_Typed_Columns_By_Name and friends), so this follows that pattern
        // rather than standing up an XlsxWorkbookWriter for data no part of this test cares about.
        // rowCount defaults to TallRowCount for the bit-packing sweep above; the memory-ceiling
        // test below scales it up instead of inventing a second fixture-generation approach.
        private static string WriteTallFixture(int rowCount = TallRowCount)
        {
            string path = Path.Combine(Path.GetTempPath(), $"excelreader-session-{Guid.NewGuid():N}.csv");
            StringBuilder csv = new("name,active,qty\n");
            for (int i = 0; i < rowCount; i++)
            {
                // Periods 3 and 5 rather than powers of two: the flags and the null pattern have to
                // stay out of phase with the 8-row byte boundary, or a batch that mis-packed its
                // final partial byte could still come out looking right.
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
                // Nullable, and blank on every fifth row, so this column's validity bitmap actually
                // has zero bits to mis-pack at a batch boundary.
                new() { Names = ["qty"], Type = NativeColumnType.Int64, Nullable = true },
            ];
        }

        private static List<TallRow> ReadTallInBatches(string path, long maxRows, out int batchCount)
        {
            Assert.Equal(NativeStatus.Ok, NativeApiTests.OpenPath(path, NativeFormat.Csv, out NativeHandle? handle));
            using NativeHandle live = handle!;
            Assert.Equal(NativeStatus.Ok,
                NativeApi.TypedParseSession.Open(live, TallSpecs(), headerRow: 1, maxRows, out NativeApi.TypedParseSession? session));
            using NativeApi.TypedParseSession open = session!;

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
                    NativeApi.FreeTable(ref table);
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

        // Same LSB-first unpacking NativeApiTests uses: bit i of byte i/8, 1 = valid, and a NULL
        // pointer means the column has no nulls at all.
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
                // Guards the fixture itself: with no nulls and no false flags, the bool values and the
                // validity bitmap would be uniform and could not expose a packing bug.
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

        // The feature's actual claim is about peak memory, not cumulative allocation - see
        // ChunkedParseBenchmark's class comment for why a BenchmarkDotNet [MemoryDiagnoser] run
        // cannot show this. Bytes owned by the single largest live NativeTable is deterministic, and
        // unlike a managed allocation count it includes the Marshal.AllocHGlobal blocks that are the
        // bulk of what this feature bounds.
        //
        // It is the OUTPUT-BLOCK half of the two things TypedParseSession's remarks promise to bound
        // ("one batch's columns plus one batch's output block"). The builder-side half needs no
        // measurement: a ColumnBuilder's ChunkedBuffer chain only ever grows by appending a new chunk
        // per row appended (ChunkedBuffer.Grow), the builders are constructed fresh inside NextBatch,
        // and nothing outside that call can reach them - so a batch's builders are bounded by the same
        // maxRows the block below is measured against, structurally rather than by assertion. If
        // NextBatch ever hoisted its builders out of the row loop, that is what this test would NOT
        // catch, and the sizes below would still pass.
        private const int CeilingRowCount = 4000;
        private const long CeilingBatchSize = 200;

        // NativeColumn's own doc comment (NativeTypedTable.cs:53-60) is the authority here: for a
        // string column, Values is the ONE allocation the column owns - ColumnBuilder.BuildStringColumn
        // (NativeApi.Typed.cs:468-478) writes the Length+1 int32 offsets array immediately followed
        // by the DataLen-byte UTF-8 blob into that single block, and Data is an interior pointer into
        // it (freeing it separately would be a double free). So (Length + 1) * sizeof(int) + DataLen
        // is the size of that one block, not two summed allocations. Every fixed-width type is Length
        // elements of its own size; Validity, when present, is a second, independent allocation of one
        // bit per row rounded up to a byte.
        private static long ColumnBytes(NativeColumn column)
        {
            long bytes = column.Type switch
            {
                NativeColumnType.String => (column.Length + 1) * sizeof(int) + column.DataLen,
                NativeColumnType.Bool => column.Length,
                NativeColumnType.Date => column.Length * sizeof(int),
                NativeColumnType.Float64 => column.Length * sizeof(double),
                _ => column.Length * sizeof(long), // Int64, Time, Timestamp
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

        // Returns the byte size of every batch NextBatch produces at the given maxRows, freeing
        // each one immediately after measuring it - the same discipline a real streaming consumer
        // follows, so the measurement cannot pass by accident from a table kept alive past its batch.
        private static List<long> MeasureBatchByteSizes(string path, long maxRows)
        {
            Assert.Equal(NativeStatus.Ok, NativeApiTests.OpenPath(path, NativeFormat.Csv, out NativeHandle? handle));
            using NativeHandle live = handle!;
            Assert.Equal(NativeStatus.Ok,
                NativeApi.TypedParseSession.Open(live, TallSpecs(), headerRow: 1, maxRows, out NativeApi.TypedParseSession? session));
            using NativeApi.TypedParseSession open = session!;

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
                    NativeApi.FreeTable(ref table);
                }
            }
            return sizes;
        }

        // The deterministic proof ChunkedParseBenchmark's class comment points to: at 200-row
        // batches over a 4000-row sheet, the largest live table must stay close to a twentieth of
        // the unbounded table's size, not equal to it. This is the test that actually exercises
        // TypedParseSession's memory-ceiling claim - the benchmark measures a different quantity
        // (cumulative managed churn) that does not and cannot show this.
        [Fact]
        public void NextBatch_Should_Bound_The_Largest_Live_Table_To_Roughly_One_Batch()
        {
            string path = WriteTallFixture(CeilingRowCount);
            try
            {
                List<long> unbounded = MeasureBatchByteSizes(path, maxRows: 0);
                List<long> batched = MeasureBatchByteSizes(path, CeilingBatchSize);

                Assert.Single(unbounded);
                // 4000 / 200 = 20 batches exactly, so this also incidentally guards against a
                // NextBatch that silently coalesced everything into one call.
                Assert.True(batched.Count > 1, $"expected more than one batch, got {batched.Count}.");

                long unboundedBytes = unbounded[0];
                long maxBatchBytes = batched.Max();

                // Generous (3x) slack absorbs fixed per-column overhead (each string column's +1
                // offset element, a validity byte shared unevenly across a batch boundary) at this
                // 200-row batch size. It is tight against the failure modes that matter most - batching
                // removed entirely, or the batch size inflated by roughly an order of magnitude, both
                // land far outside 3x the ideal ~5% fraction - but it is not tight in general: a
                // partial regression that merely doubled or tripled the intended batch size would
                // still pass. Verified against the total-removal case by temporarily changing
                // NextBatch's row loop to ignore maxRows and confirming this assertion fails.
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

            int status = NativeApi.OpenTypedReader(live, Specs(), headerRow: 0, maxRows: -1, out nint reader);

            Assert.Equal(NativeStatus.InvalidArgument, status);
            Assert.Equal(0, reader);
        }

        [Fact]
        public void OpenTypedReader_Should_Reject_A_Null_Handle()
        {
            Assert.Equal(NativeStatus.InvalidHandle,
                NativeApi.OpenTypedReader(null, Specs(), headerRow: 0, maxRows: 10, out nint reader));
            Assert.Equal(0, reader);
        }

        [Fact]
        public void NextTypedBatch_Should_Report_InvalidHandle_For_A_Closed_Reader()
        {
            Assert.Equal(NativeStatus.Ok, NativeApiTests.OpenPath(XlsxFixture, NativeFormat.Auto, out NativeHandle? handle));
            using NativeHandle live = handle!;
            Assert.Equal(NativeStatus.Ok, NativeApi.OpenTypedReader(live, Specs(), 0, 10, out nint reader));

            NativeApi.CloseTypedReader(reader);

            Assert.Equal(NativeStatus.InvalidHandle, NativeApi.NextTypedBatch(reader, out NativeTable table));
            Assert.Equal(IntPtr.Zero, table.Columns);
        }

        [Fact]
        public void CloseTypedReader_Should_Be_A_No_Op_On_Zero_And_On_A_Second_Close()
        {
            Assert.Equal(NativeStatus.Ok, NativeApiTests.OpenPath(XlsxFixture, NativeFormat.Auto, out NativeHandle? handle));
            using NativeHandle live = handle!;
            Assert.Equal(NativeStatus.Ok, NativeApi.OpenTypedReader(live, Specs(), 0, 10, out nint reader));

            NativeApi.CloseTypedReader(reader);
            NativeApi.CloseTypedReader(reader); // must not throw
            NativeApi.CloseTypedReader(0);      // must not throw
        }

        // A reader id must never resolve to a workbook id, and vice versa - NativeHandleTable's
        // type check is what makes a cross-kind call a clean no-op instead of a wrong-object free.
        [Fact]
        public void NextTypedBatch_Should_Reject_A_Workbook_Id()
        {
            Assert.Equal(NativeStatus.Ok, NativeApiTests.OpenPath(XlsxFixture, NativeFormat.Auto, out NativeHandle? handle));
            using NativeHandle live = handle!;
            nint workbookId = NativeHandleTable.Register(live);

            Assert.Equal(NativeStatus.InvalidHandle, NativeApi.NextTypedBatch(workbookId, out _));
        }
    }
}
