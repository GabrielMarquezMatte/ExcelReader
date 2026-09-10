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

        // Enough rows that a 10-row batch leaves the overwhelming majority unread: a session that
        // resumed after an interleaved call would come back with data, which is exactly the failure
        // these tests have to be able to see.
        private const int SmallRowCount = 500;

        // The scenario the interlock exists for, at the size it was measured at: 200,000 rows read in
        // 1,000-row batches, with one ordinary xl_parse_typed interleaved after the first batch. Before
        // the interlock that call rewound the shared CSV stream under the live reader, which then
        // reported XL_EOF - a clean success, with an EMPTY xl_last_error - after 12,774 of its 200,000
        // rows: 94% of the data dropped with no way for the caller to notice. Kept at full size
        // because a short fixture cannot show the difference between truncation and completion; a
        // 200k-row single-column CSV parses in well under a second.
        private const int InterleaveRowCount = 200_000;
        private const long InterleaveBatchSize = 1000;

        private static NativeColumnSpec[] IdSpecs()
        {
            return [new() { Names = ["id"], Type = NativeColumnType.Int64 }];
        }

        // Index-based, so it needs no header row - what the header-less sample.xlsx opens are for.
        private static NativeColumnSpec[] FirstColumnSpecs()
        {
            return [new() { Index = 0, Type = NativeColumnType.String, Nullable = true }];
        }

        // One non-nullable int64 column: the cheapest fixture that can also carry an unconvertible
        // value (see WriteUnconvertibleCsv) for the conversion-fault tests.
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

        // Row 3 cannot convert to int64 and the column is non-nullable, so the fault lands on the
        // SECOND batch at a batch size of 2 - after one batch has already been handed out and freed,
        // which is what makes "batches already handed out stay valid" observable.
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

        // ---- one live session per workbook ---------------------------------------------------

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

            // Closing the first releases the slot, so the refusal is a live-session rule and not a
            // once-per-workbook rule.
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

            // The stream's own release frees the slot the same way xl_typed_reader_close does.
            ReleaseStream(ref stream);
            Assert.Equal(NativeStatus.Ok, NativeApi.OpenTypedReader(live, FirstColumnSpecs(), headerRow: 0, maxRows: 1, out nint after));
            NativeApi.CloseTypedReader(after);
        }

        // ---- the measured interleaving scenario ----------------------------------------------

        // See InterleaveRowCount for the measurement this pins. The reader must fail loudly rather
        // than resume from a rewound cursor, and xl_parse_typed's own result must still be complete -
        // the interleaved call is the one entitled to the workbook for the length of its own call.
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
                    // The interleaved call is the one that is allowed to succeed - it owns the
                    // workbook's cursor for the duration of its own call.
                    Assert.Equal((long)InterleaveRowCount, interleaved);

                    // And the reader is dead, loudly: not XL_EOF, not a short batch.
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

        // ---- every other read on the workbook faults a live reader ----------------------------

        // Shared by the per-entry-point tests below: one live reader, one batch already taken, then
        // `interleaved` runs on the same workbook. The reader must report XL_ERROR with a message,
        // and must keep reporting the same one - a faulted reader never resumes.
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
            // Span<byte>.Empty asks for the row's size, which is what actually advances the
            // workbook's own cursor - the copy-out afterwards is irrelevant here.
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

        // excelreader.h's chunked-reading section promises exactly this. Index 0 is deliberate: the
        // rule is not "moving to a DIFFERENT sheet", it is that xl_move_to_sheet drops and rebuilds
        // the workbook's row cursor, which invalidates the reader's position either way.
        [Fact]
        public void MoveToSheet_Should_Fault_A_Live_Reader()
        {
            AssertFaultsTheLiveReader(static live => Assert.Equal(NativeStatus.Ok, NativeApi.MoveToSheet(live, 0)));
        }

        // ---- closing the workbook mid-session -------------------------------------------------

        // The spec's Testing section, and the other half of the header's borrow rule: xl_close while
        // a reader is open is a clean XL_ERROR, not a crash, and it latches.
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

                    live.Dispose(); // xl_close

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

        // ---- the conversion fault itself ------------------------------------------------------

        // The LATCHES promise in excelreader.h, exercised through the status codes rather than
        // inferred from the implementation: a non-nullable column whose value fails to convert is
        // XL_ERROR, with a message, on this call and on every call after it.
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
                    // Rows 1 and 2 convert, so the first batch is a normal success the caller owns.
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
                    // Same message, not an empty one: a caller that only reads xl_last_error after the
                    // second call still learns what went wrong.
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

        // xl_parse_typed opens a session of its own internally. It must not be refused by, or trip
        // over, the one-live-session rule - it drains and closes inside its own call, so two calls in
        // a row have to be indistinguishable from one.
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

        // The Arrow release callbacks are plain function pointers in the structs, exactly as a C
        // consumer sees them - same invocation shape ArrowStreamTests uses.
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
