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
            List<string> batched = ReadInBatches(maxRows, out _);

            Assert.Equal(1, wholeBatches);
            Assert.Equal(whole, batched);
        }
    }
}
