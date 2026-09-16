using System.Globalization;
using System.Runtime.InteropServices;
using System.Text;
using ExcelReader.Native;

namespace ExcelReader.Tests
{
    /// <summary>
    /// The one-live-session rule, from the outside. A <c>xl_typed_reader</c>/Arrow stream is the only
    /// thing in this ABI that holds an <c>IExcelRowEnumerator</c> open ACROSS calls, and
    /// <see cref="Core.Reader.IExcelRowReader"/> serves one usable enumerator at a time — so every
    /// other read on the same workbook either has to be refused up front or has to invalidate the
    /// live session loudly. These tests pin both halves; without the interlock they fail by returning
    /// XL_OK/XL_EOF with silently truncated data.
    /// </summary>
    public sealed class SessionLifetimeTests
    {
        private static readonly string XlsxFixture = Path.Combine(AppContext.BaseDirectory, "data", "sample.xlsx");

        private const int SmallRowCount = 500;

        private const int InterleaveRowCount = 200_000;
        private const long InterleaveBatchSize = 1000;

        private static NativeColumnSpec[] IdSpecs()
        {
            return [new() { Names = ["id"], Type = NativeColumnType.Int64 }];
        }

        private static NativeColumnSpec[] FirstColumnSpecs()
        {
            return [new() { Index = 0, Type = NativeColumnType.String, Nullable = true }];
        }

        private static string WriteCsv(int rowCount)
        {
            string path = Path.Combine(Path.GetTempPath(), $"excelreader-lifetime-{Guid.NewGuid():N}.csv");
            StringBuilder csv = new("id\n");
            for (int i = 0; i < rowCount; i++)
            {
                csv.Append(CultureInfo.InvariantCulture, $"{i}\n");
            }
            File.WriteAllText(path, csv.ToString());
            return path;
        }

        private static string WriteUnconvertibleCsv()
        {
            string path = Path.Combine(Path.GetTempPath(), $"excelreader-lifetime-{Guid.NewGuid():N}.csv");
            File.WriteAllText(path, "id\n1\n2\nnot-a-number\n4\n");
            return path;
        }

        private static long RowsIn(NativeTable table)
        {
            return table.RowCount;
        }


        [Fact]
        public void OpenTypedReader_Should_Refuse_A_Second_Live_Reader_On_The_Same_Workbook()
        {
            Assert.Equal(NativeStatus.Ok, NativeApiTests.OpenPath(XlsxFixture, NativeFormat.Auto, out NativeHandle? handle));
            using NativeHandle live = handle!;
            Assert.Equal(NativeStatus.Ok, NativeApi.OpenTypedReader(live, FirstColumnSpecs(), headerRow: 0, maxRows: 1, out nint first));

            int status = NativeApi.OpenTypedReader(live, FirstColumnSpecs(), headerRow: 0, maxRows: 1, out nint second);

            Assert.Equal(NativeStatus.Error, status);
            Assert.Equal(0, second);
            Assert.NotEmpty(NativeApi.LastErrorText());

            NativeApi.CloseTypedReader(first);
            Assert.Equal(NativeStatus.Ok, NativeApi.OpenTypedReader(live, FirstColumnSpecs(), headerRow: 0, maxRows: 1, out nint third));
            NativeApi.CloseTypedReader(third);
        }

        [Fact]
        public void OpenArrowStream_Should_Refuse_A_Stream_While_A_Typed_Reader_Is_Open()
        {
            Assert.Equal(NativeStatus.Ok, NativeApiTests.OpenPath(XlsxFixture, NativeFormat.Auto, out NativeHandle? handle));
            using NativeHandle live = handle!;
            Assert.Equal(NativeStatus.Ok, NativeApi.OpenTypedReader(live, FirstColumnSpecs(), headerRow: 0, maxRows: 1, out nint reader));

            int status = NativeApi.OpenArrowStream(live, FirstColumnSpecs(), headerRow: 0, maxRows: 1, out ArrowArrayStream stream);

            Assert.Equal(NativeStatus.Error, status);
            Assert.Equal(IntPtr.Zero, stream.Release);
            Assert.NotEmpty(NativeApi.LastErrorText());
            NativeApi.CloseTypedReader(reader);
        }

        [Fact]
        public void OpenTypedReader_Should_Refuse_A_Reader_While_An_Arrow_Stream_Is_Open()
        {
            Assert.Equal(NativeStatus.Ok, NativeApiTests.OpenPath(XlsxFixture, NativeFormat.Auto, out NativeHandle? handle));
            using NativeHandle live = handle!;
            Assert.Equal(NativeStatus.Ok,
                NativeApi.OpenArrowStream(live, FirstColumnSpecs(), headerRow: 0, maxRows: 1, out ArrowArrayStream stream));

            int status = NativeApi.OpenTypedReader(live, FirstColumnSpecs(), headerRow: 0, maxRows: 1, out nint reader);

            Assert.Equal(NativeStatus.Error, status);
            Assert.Equal(0, reader);

            ReleaseStream(ref stream);
            Assert.Equal(NativeStatus.Ok, NativeApi.OpenTypedReader(live, FirstColumnSpecs(), headerRow: 0, maxRows: 1, out nint after));
            NativeApi.CloseTypedReader(after);
        }


        [Fact]
        public void ParseTyped_Should_Fault_A_Live_Reader_Instead_Of_Truncating_It()
        {
            string path = WriteCsv(InterleaveRowCount);
            try
            {
                Assert.Equal(NativeStatus.Ok, NativeApiTests.OpenPath(path, NativeFormat.Csv, out NativeHandle? handle));
                using NativeHandle live = handle!;
                Assert.Equal(NativeStatus.Ok,
                    NativeApi.OpenTypedReader(live, IdSpecs(), headerRow: 1, InterleaveBatchSize, out nint reader));
                try
                {
                    Assert.Equal(NativeStatus.Ok, NativeApi.NextTypedBatch(reader, out NativeTable batch));
                    long fromReader = RowsIn(batch);
                    NativeApi.FreeTable(ref batch);
                    Assert.Equal(InterleaveBatchSize, fromReader);

                    Assert.Equal(NativeStatus.Ok, NativeApi.ParseTyped(live, IdSpecs(), headerRow: 1, out NativeTable whole));
                    long interleaved = RowsIn(whole);
                    NativeApi.FreeTable(ref whole);
                    Assert.Equal((long)InterleaveRowCount, interleaved);

                    Assert.Equal(NativeStatus.Error, NativeApi.NextTypedBatch(reader, out NativeTable after));
                    Assert.Equal(IntPtr.Zero, after.Columns);
                    Assert.NotEmpty(NativeApi.LastErrorText());
                }
                finally
                {
                    NativeApi.CloseTypedReader(reader);
                }
            }
            finally
            {
                File.Delete(path);
            }
        }


        private static void AssertFaultsTheLiveReader(Action<NativeHandle> interleaved)
        {
            string path = WriteCsv(SmallRowCount);
            try
            {
                Assert.Equal(NativeStatus.Ok, NativeApiTests.OpenPath(path, NativeFormat.Csv, out NativeHandle? handle));
                using NativeHandle live = handle!;
                Assert.Equal(NativeStatus.Ok,
                    NativeApi.OpenTypedReader(live, IdSpecs(), headerRow: 1, maxRows: 10, out nint reader));
                try
                {
                    Assert.Equal(NativeStatus.Ok, NativeApi.NextTypedBatch(reader, out NativeTable first));
                    NativeApi.FreeTable(ref first);

                    interleaved(live);

                    Assert.Equal(NativeStatus.Error, NativeApi.NextTypedBatch(reader, out NativeTable after));
                    Assert.Equal(IntPtr.Zero, after.Columns);
                    string latched = NativeApi.LastErrorText();
                    Assert.NotEmpty(latched);

                    Assert.Equal(NativeStatus.Error, NativeApi.NextTypedBatch(reader, out NativeTable again));
                    Assert.Equal(IntPtr.Zero, again.Columns);
                    Assert.Equal(latched, NativeApi.LastErrorText());
                }
                finally
                {
                    NativeApi.CloseTypedReader(reader);
                }
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void ParseArrow_Should_Fault_A_Live_Reader()
        {
            AssertFaultsTheLiveReader(static live =>
            {
                Assert.Equal(NativeStatus.Ok,
                    NativeApi.ParseArrow(live, IdSpecs(), headerRow: 1, out ArrowArray array, out ArrowSchema schema));
                ReleaseArray(ref array);
                ReleaseSchema(ref schema);
            });
        }

        [Fact]
        public void NextRow_Should_Fault_A_Live_Reader()
        {
            AssertFaultsTheLiveReader(static live =>
                Assert.Equal(NativeStatus.BufferTooSmall, NativeApi.NextRow(live, Span<byte>.Empty, out _)));
        }

        [Fact]
        public void ReadAllBlob_Should_Fault_A_Live_Reader()
        {
            AssertFaultsTheLiveReader(static live =>
                Assert.Equal(NativeStatus.BufferTooSmall, NativeApi.ReadAllBlob(live, Span<byte>.Empty, out _)));
        }

        [Fact]
        public void ReadAllDecoded_Should_Fault_A_Live_Reader()
        {
            AssertFaultsTheLiveReader(static live =>
            {
                Assert.Equal(NativeStatus.Ok, NativeApi.ReadAllDecoded(live, out NativeRows rows));
                NativeApi.FreeRows(ref rows);
            });
        }

        [Fact]
        public void InferSchema_Should_Fault_A_Live_Reader()
        {
            AssertFaultsTheLiveReader(static live =>
            {
                Assert.Equal(NativeStatus.Ok, NativeApi.InferSchema(live, headerRow: 1, sampleSize: 8, out NativeInferredSchema schema));
                NativeApi.FreeSchema(ref schema);
            });
        }

        [Fact]
        public void MoveToSheet_Should_Fault_A_Live_Reader()
        {
            AssertFaultsTheLiveReader(static live => Assert.Equal(NativeStatus.Ok, NativeApi.MoveToSheet(live, 0)));
        }


        [Fact]
        public void Dispose_Should_Fault_A_Live_Reader_And_Latch()
        {
            string path = WriteCsv(SmallRowCount);
            try
            {
                Assert.Equal(NativeStatus.Ok, NativeApiTests.OpenPath(path, NativeFormat.Csv, out NativeHandle? handle));
                NativeHandle live = handle!;
                Assert.Equal(NativeStatus.Ok,
                    NativeApi.OpenTypedReader(live, IdSpecs(), headerRow: 1, maxRows: 10, out nint reader));
                try
                {
                    Assert.Equal(NativeStatus.Ok, NativeApi.NextTypedBatch(reader, out NativeTable first));
                    NativeApi.FreeTable(ref first);

                    live.Dispose();

                    Assert.Equal(NativeStatus.Error, NativeApi.NextTypedBatch(reader, out NativeTable after));
                    Assert.Equal(IntPtr.Zero, after.Columns);
                    string latched = NativeApi.LastErrorText();
                    Assert.NotEmpty(latched);

                    Assert.Equal(NativeStatus.Error, NativeApi.NextTypedBatch(reader, out _));
                    Assert.Equal(latched, NativeApi.LastErrorText());
                }
                finally
                {
                    NativeApi.CloseTypedReader(reader);
                }
            }
            finally
            {
                File.Delete(path);
            }
        }


        [Fact]
        public void ParseTyped_Should_Not_Fault_A_Live_Reader_When_Argument_Validation_Fails()
        {
            string path = WriteCsv(SmallRowCount);
            try
            {
                Assert.Equal(NativeStatus.Ok, NativeApiTests.OpenPath(path, NativeFormat.Csv, out NativeHandle? handle));
                using NativeHandle live = handle!;
                Assert.Equal(NativeStatus.Ok,
                    NativeApi.OpenTypedReader(live, IdSpecs(), headerRow: 1, maxRows: 10, out nint reader));
                try
                {
                    Assert.Equal(NativeStatus.Ok, NativeApi.NextTypedBatch(reader, out NativeTable first));
                    NativeApi.FreeTable(ref first);

                    NativeColumnSpec[] blankNameSpec = [new() { Names = [" "], Type = NativeColumnType.Int64 }];
                    int status = NativeApi.ParseTyped(live, blankNameSpec, headerRow: 1, out NativeTable invalid);

                    Assert.Equal(NativeStatus.InvalidArgument, status);
                    Assert.Equal(IntPtr.Zero, invalid.Columns);
                    Assert.NotEmpty(NativeApi.LastErrorText());

                    Assert.Equal(NativeStatus.Ok, NativeApi.NextTypedBatch(reader, out NativeTable after));
                    long rowsAfter = RowsIn(after);
                    NativeApi.FreeTable(ref after);
                    Assert.True(rowsAfter > 0);
                }
                finally
                {
                    NativeApi.CloseTypedReader(reader);
                }
            }
            finally
            {
                File.Delete(path);
            }
        }


        [Fact]
        public void NextTypedBatch_Should_Latch_A_Conversion_Failure()
        {
            string path = WriteUnconvertibleCsv();
            try
            {
                Assert.Equal(NativeStatus.Ok, NativeApiTests.OpenPath(path, NativeFormat.Csv, out NativeHandle? handle));
                using NativeHandle live = handle!;
                Assert.Equal(NativeStatus.Ok,
                    NativeApi.OpenTypedReader(live, IdSpecs(), headerRow: 1, maxRows: 2, out nint reader));
                try
                {
                    Assert.Equal(NativeStatus.Ok, NativeApi.NextTypedBatch(reader, out NativeTable good));
                    Assert.Equal(2L, RowsIn(good));
                    NativeApi.FreeTable(ref good);

                    Assert.Equal(NativeStatus.Error, NativeApi.NextTypedBatch(reader, out NativeTable bad));
                    Assert.Equal(IntPtr.Zero, bad.Columns);
                    string latched = NativeApi.LastErrorText();
                    Assert.NotEmpty(latched);
                    Assert.Contains("failed to convert", latched, StringComparison.Ordinal);

                    Assert.Equal(NativeStatus.Error, NativeApi.NextTypedBatch(reader, out NativeTable again));
                    Assert.Equal(IntPtr.Zero, again.Columns);
                    Assert.Equal(latched, NativeApi.LastErrorText());
                }
                finally
                {
                    NativeApi.CloseTypedReader(reader);
                }
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void ParseTyped_Should_Stay_Repeatable_On_The_Same_Workbook()
        {
            string path = WriteCsv(SmallRowCount);
            try
            {
                Assert.Equal(NativeStatus.Ok, NativeApiTests.OpenPath(path, NativeFormat.Csv, out NativeHandle? handle));
                using NativeHandle live = handle!;

                Assert.Equal(NativeStatus.Ok, NativeApi.ParseTyped(live, IdSpecs(), headerRow: 1, out NativeTable first));
                long firstRows = RowsIn(first);
                NativeApi.FreeTable(ref first);
                Assert.Equal(NativeStatus.Ok, NativeApi.ParseTyped(live, IdSpecs(), headerRow: 1, out NativeTable second));
                long secondRows = RowsIn(second);
                NativeApi.FreeTable(ref second);

                Assert.Equal((long)SmallRowCount, firstRows);
                Assert.Equal(firstRows, secondRows);
            }
            finally
            {
                File.Delete(path);
            }
        }

        private delegate void ReleaseStreamFn(ref ArrowArrayStream stream);
        private delegate void ReleaseArrayFn(ref ArrowArray array);
        private delegate void ReleaseSchemaFn(ref ArrowSchema schema);

        private static void ReleaseStream(ref ArrowArrayStream stream)
        {
            Marshal.GetDelegateForFunctionPointer<ReleaseStreamFn>(stream.Release)(ref stream);
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
