using System.Globalization;
using System.Runtime.InteropServices;
using System.Text;
using ExcelReader.Native;

namespace ExcelReader.Tests
{
    public sealed class ArrowStreamTests
    {
        private static readonly string XlsxFixture = Path.Combine(AppContext.BaseDirectory, "data", "sample.xlsx");

        private static NativeColumnSpec[] Specs()
        {
            return [new() { Index = 0, Type = NativeColumnType.String, Nullable = true }];
        }

        private static ArrowArrayStream Open(long maxRows, out NativeHandle live)
        {
            Assert.Equal(NativeStatus.Ok, NativeApiTests.OpenPath(XlsxFixture, NativeFormat.Auto, out NativeHandle? handle));
            live = handle!;
            Assert.Equal(NativeStatus.Ok,
                NativeApi.OpenArrowStream(live, Specs(), headerRow: 0, maxRows, out ArrowArrayStream stream));
            return stream;
        }

        // Arrow's contract: end-of-stream is get_next returning 0 with a RELEASED array, never an
        // error code. Getting this wrong breaks every consumer subtly, so it is pinned directly.
        [Fact]
        public void GetNext_Should_Signal_End_Of_Stream_With_Zero_And_A_Null_Release()
        {
            ArrowArrayStream stream = Open(maxRows: 1, out NativeHandle live);
            using (live)
            {
                try
                {
                    int batches = 0;
                    while (true)
                    {
                        Assert.Equal(0, InvokeGetNext(ref stream, out ArrowArray array));
                        if (array.Release == IntPtr.Zero)
                        {
                            break;
                        }
                        batches++;
                        ReleaseArray(ref array);
                    }
                    Assert.True(batches > 1, "a batch size of 1 must produce more than one batch");
                }
                finally
                {
                    InvokeRelease(ref stream);
                }
            }
        }

        // One child schema flattened to the three things an Arrow consumer actually validates. A
        // child count alone would let a stream whose format codes, names or nullability drifted
        // between batches pass, which is precisely the breakage the spec's same-schema rule forbids.
        private readonly record struct SchemaShape(string? Format, string? Name, long Flags);

        private static List<SchemaShape> DescribeChildren(ArrowSchema schema)
        {
            List<SchemaShape> children = [];
            for (long i = 0; i < schema.NChildren; i++)
            {
                ArrowSchema child = Marshal.PtrToStructure<ArrowSchema>(
                    Marshal.ReadIntPtr(schema.Children, (int)(i * IntPtr.Size)));
                children.Add(new SchemaShape(
                    Marshal.PtrToStringUTF8(child.Format),
                    Marshal.PtrToStringUTF8(child.Name),
                    child.Flags));
            }
            return children;
        }

        // The Arrow spec requires every batch in a stream to carry the same schema.
        [Fact]
        public void GetSchema_Should_Return_The_Same_Shape_Every_Time()
        {
            ArrowArrayStream stream = Open(maxRows: 1, out NativeHandle live);
            using (live)
            {
                try
                {
                    Assert.Equal(0, InvokeGetSchema(ref stream, out ArrowSchema first));
                    Assert.Equal(0, InvokeGetSchema(ref stream, out ArrowSchema second));

                    List<SchemaShape> firstChildren = DescribeChildren(first);
                    Assert.Equal(firstChildren, DescribeChildren(second));
                    Assert.Equal("+s", Marshal.PtrToStringUTF8(first.Format));
                    Assert.Equal(Marshal.PtrToStringUTF8(first.Format), Marshal.PtrToStringUTF8(second.Format));

                    // Ground truth for Specs(), so this cannot pass by both calls agreeing on the
                    // wrong answer: one unnamed nullable string column, named by its index.
                    Assert.Equal([new SchemaShape("u", "0", ArrowFlags.Nullable)], firstChildren);

                    ReleaseSchema(ref first);
                    ReleaseSchema(ref second);
                }
                finally
                {
                    InvokeRelease(ref stream);
                }
            }
        }

        // Abandoning a stream part-way must not leak the session or the enumerator.
        [Fact]
        public void Release_Should_Be_Safe_Mid_Stream_And_Idempotent()
        {
            ArrowArrayStream stream = Open(maxRows: 1, out NativeHandle live);
            using (live)
            {
                Assert.Equal(0, InvokeGetNext(ref stream, out ArrowArray array));
                Assert.NotEqual(IntPtr.Zero, array.Release);
                ReleaseArray(ref array);

                InvokeRelease(ref stream);
                Assert.Equal(IntPtr.Zero, stream.Release);
                InvokeRelease(ref stream); // must not throw

                // The stream BORROWS the workbook - release closes the read, never the handle. If it
                // ever disposed the workbook instead, this second open would fail rather than hand
                // back a fresh stream over the same still-live handle.
                Assert.Equal(NativeStatus.Ok,
                    NativeApi.OpenArrowStream(live, Specs(), headerRow: 0, maxRows: 1, out ArrowArrayStream reopened));
                Assert.Equal(0, InvokeGetNext(ref reopened, out ArrowArray fromReopened));
                Assert.NotEqual(IntPtr.Zero, fromReopened.Release);
                ReleaseArray(ref fromReopened);
                InvokeRelease(ref reopened);
            }
        }

        // Arrow's boolean layout and its validity bitmap are both LSB-first bit-packed, and the repack
        // from xl_column's byte-per-row bools (NativeApi.Arrow.cs's BitPackBoolColumn) runs ONLY on this
        // export path - the NativeTable batching sweep in TypedParseSessionTests never reaches it. A
        // batch whose row count is not a multiple of 8 ends mid-byte, which is exactly where an
        // LSB-first packing bug hides, so this drives the stream at such sizes and compares against the
        // single unbounded batch.
        private const int TallRowCount = 43;

        // One decoded row, compared as a whole so a mis-packed bool bit or validity bit fails the
        // assertion rather than being averaged away.
        private readonly record struct TallRow(bool Active, bool QtyValid, long Qty);

        // Same temp-CSV approach TypedParseSessionTests uses for its own tall fixture: 43 rows is
        // coprime with every batch size below, and the flag/null periods (3 and 5) stay out of phase
        // with the 8-row byte boundary so a mis-packed final partial byte cannot look correct anyway.
        private static string WriteTallFixture()
        {
            string path = Path.Combine(Path.GetTempPath(), $"excelreader-arrowstream-{Guid.NewGuid():N}.csv");
            StringBuilder csv = new("active,qty\n");
            for (int i = 0; i < TallRowCount; i++)
            {
                string active = i % 3 == 0 ? "true" : "false";
                string qty = i % 5 == 0 ? string.Empty : (i * 3).ToString(CultureInfo.InvariantCulture);
                csv.Append(CultureInfo.InvariantCulture, $"{active},{qty}\n");
            }
            File.WriteAllText(path, csv.ToString());
            return path;
        }

        private static NativeColumnSpec[] TallSpecs()
        {
            return
            [
                new() { Names = ["active"], Type = NativeColumnType.Bool },
                // Nullable, and blank on every fifth row, so this column's validity bitmap actually has
                // zero bits to mis-pack at a batch boundary.
                new() { Names = ["qty"], Type = NativeColumnType.Int64, Nullable = true },
            ];
        }

        private static List<TallRow> ReadTallThroughStream(string path, long maxRows, out int batchCount)
        {
            Assert.Equal(NativeStatus.Ok, NativeApiTests.OpenPath(path, NativeFormat.Csv, out NativeHandle? handle));
            using NativeHandle live = handle!;
            Assert.Equal(NativeStatus.Ok,
                NativeApi.OpenArrowStream(live, TallSpecs(), headerRow: 1, maxRows, out ArrowArrayStream stream));

            List<TallRow> rows = [];
            batchCount = 0;
            try
            {
                while (true)
                {
                    Assert.Equal(0, InvokeGetNext(ref stream, out ArrowArray array));
                    if (array.Release == IntPtr.Zero)
                    {
                        break;
                    }
                    batchCount++;
                    try
                    {
                        rows.AddRange(DecodeTallBatch(array));
                    }
                    finally
                    {
                        ReleaseArray(ref array);
                    }
                }
            }
            finally
            {
                InvokeRelease(ref stream);
            }
            return rows;
        }

        private static List<TallRow> DecodeTallBatch(ArrowArray array)
        {
            Assert.Equal(2, array.NChildren);
            ArrowArray active = ChildAt(array, 0);
            ArrowArray qty = ChildAt(array, 1);
            int rowCount = (int)array.Length;

            // Buffer 0 is always validity; the bool column's values live in buffer 1 bit-packed, the
            // int64 column's in buffer 1 as eight bytes per row.
            bool[] flags = DecodeBits(BufferAt(active, 1), rowCount, whenAbsent: false);
            bool[] valid = DecodeBits(BufferAt(qty, 0), rowCount, whenAbsent: true);
            long[] quantities = new long[rowCount];
            Marshal.Copy(BufferAt(qty, 1), quantities, 0, rowCount);

            // null_count must agree with the validity bitmap this same batch carries. CountUnset
            // popcounts whole bytes on the assumption that every bit past `length` is zero, so a batch
            // whose row count is not a multiple of 8 is exactly where a stale tail bit would inflate
            // null_count - a discrepancy the row values alone cannot show, and one pyarrow would act on
            // because it trusts null_count for its validity fast paths.
            Assert.Equal(valid.Count(v => !v), qty.NullCount);
            Assert.Equal(0, active.NullCount); // non-nullable: no validity bitmap, so no nulls
            Assert.Equal(0, array.NullCount);  // the top-level struct array has no row-level nulls

            List<TallRow> decoded = [];
            for (int i = 0; i < rowCount; i++)
            {
                decoded.Add(new TallRow(flags[i], valid[i], quantities[i]));
            }
            return decoded;
        }

        // Ground truth straight from WriteTallFixture's generating rules, so the comparison below does
        // not rest solely on the unbounded read agreeing with the batched one: both sides run the same
        // PackBitsLsbFirst, so a batch-size-invariant bug (a global bit inversion, say) would cancel
        // out between them and pass. These assertions are absolute.
        private static void AssertMatchesFixtureRules(List<TallRow> rows)
        {
            Assert.Equal(TallRowCount, rows.Count);
            for (int i = 0; i < rows.Count; i++)
            {
                Assert.Equal(i % 3 == 0, rows[i].Active);
                Assert.Equal(i % 5 != 0, rows[i].QtyValid);
                if (rows[i].QtyValid)
                {
                    Assert.Equal(i * 3, rows[i].Qty);
                }
            }
        }

        private static ArrowArray ChildAt(ArrowArray array, int index)
        {
            return Marshal.PtrToStructure<ArrowArray>(Marshal.ReadIntPtr(array.Children, index * IntPtr.Size));
        }

        private static IntPtr BufferAt(ArrowArray array, int index)
        {
            return Marshal.ReadIntPtr(array.Buffers, index * IntPtr.Size);
        }

        // LSB-first: bit i lives in bit (i & 7) of byte (i >> 3). A NULL buffer is Arrow's "absent"
        // marker, which for a validity bitmap means every row is valid.
        private static bool[] DecodeBits(IntPtr bits, int rowCount, bool whenAbsent)
        {
            bool[] result = new bool[rowCount];
            if (bits == IntPtr.Zero)
            {
                Array.Fill(result, whenAbsent);
                return result;
            }
            byte[] bitmap = new byte[(rowCount + 7) / 8];
            Marshal.Copy(bits, bitmap, 0, bitmap.Length);
            for (int i = 0; i < rowCount; i++)
            {
                result[i] = (bitmap[i >> 3] & (1 << (i & 7))) != 0;
            }
            return result;
        }

        [Theory]
        [InlineData(3)]
        [InlineData(7)]
        [InlineData(9)]
        public void GetNext_Should_Match_The_Unbounded_Batch_Across_Bit_Packing_Boundaries(long maxRows)
        {
            string path = WriteTallFixture();
            try
            {
                List<TallRow> whole = ReadTallThroughStream(path, 0, out int wholeBatches);
                List<TallRow> batched = ReadTallThroughStream(path, maxRows, out int batches);

                Assert.Equal(1, wholeBatches);
                // Guards the fixture: uniform flags or no nulls at all would leave nothing to mis-pack.
                Assert.Contains(whole, row => row.Active);
                Assert.Contains(whole, row => !row.Active);
                Assert.Contains(whole, row => row.QtyValid);
                Assert.Contains(whole, row => !row.QtyValid);

                // Absolute first, then relative: the batched read must match the fixture's own rules,
                // and the unbounded read must agree with it.
                AssertMatchesFixtureRules(batched);
                AssertMatchesFixtureRules(whole);
                Assert.Equal(whole, batched);
                long expectedBatches = (whole.Count + maxRows - 1) / maxRows;
                Assert.Equal(expectedBatches, batches);
            }
            finally
            {
                File.Delete(path);
            }
        }

        private delegate int GetNextFn(ref ArrowArrayStream stream, out ArrowArray array);
        private delegate int GetSchemaFn(ref ArrowArrayStream stream, out ArrowSchema schema);
        private delegate void ReleaseStreamFn(ref ArrowArrayStream stream);
        private delegate void ReleaseArrayFn(ref ArrowArray array);
        private delegate void ReleaseSchemaFn(ref ArrowSchema schema);

        private static int InvokeGetNext(ref ArrowArrayStream stream, out ArrowArray array)
        {
            return Marshal.GetDelegateForFunctionPointer<GetNextFn>(stream.GetNext)(ref stream, out array);
        }

        private static int InvokeGetSchema(ref ArrowArrayStream stream, out ArrowSchema schema)
        {
            return Marshal.GetDelegateForFunctionPointer<GetSchemaFn>(stream.GetSchema)(ref stream, out schema);
        }

        private static void InvokeRelease(ref ArrowArrayStream stream)
        {
            if (stream.Release != IntPtr.Zero)
            {
                Marshal.GetDelegateForFunctionPointer<ReleaseStreamFn>(stream.Release)(ref stream);
            }
        }

        private static void ReleaseArray(ref ArrowArray array)
        {
            Marshal.GetDelegateForFunctionPointer<ReleaseArrayFn>(array.Release)(ref array);
        }

        private static void ReleaseSchema(ref ArrowSchema schema)
        {
            Marshal.GetDelegateForFunctionPointer<ReleaseSchemaFn>(schema.Release)(ref schema);
        }
    }
}
