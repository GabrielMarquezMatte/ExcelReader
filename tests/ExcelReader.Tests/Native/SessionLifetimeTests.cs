using System.Globalization;
using System.Runtime.InteropServices;
using System.Text;
using ExcelReader.Native;
using ExcelReader.Native.Arrow;
using ExcelReader.Native.Reading;
using ExcelReader.Native.Typed;

namespace ExcelReader.Tests.Native
{
    /// <summary>
    /// A typed reader and an Arrow stream each own their enumerator, so reads on one workbook do not
    /// disturb one another and outlive the workbook handle.
    /// </summary>
    public sealed class SessionLifetimeTests
    {
        private static readonly string XlsxFixture = Path.Combine(AppContext.BaseDirectory, "data", "sample.xlsx");

        private const int SmallRowCount = 500;

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


        private static long Drain(nint reader)
        {
            long total = 0;
            while (true)
            {
                int status = TypedApi.NextTypedBatch(reader, out NativeTable batch);
                if (status == NativeStatus.Eof)
                {
                    return total;
                }
                Assert.Equal(NativeStatus.Ok, status);
                total += batch.RowCount;
                TypedApi.FreeTable(ref batch);
            }
        }

        [Fact]
        public void Two_Typed_Readers_On_One_Workbook_Should_Both_Read_Every_Row()
        {
            string path = WriteCsv(SmallRowCount);
            try
            {
                Assert.Equal(NativeStatus.Ok, NativeApiTests.OpenPath(path, NativeFormat.Csv, out NativeHandle? handle));
                using NativeHandle live = handle!;
                Assert.Equal(NativeStatus.Ok, TypedApi.OpenTypedReader(live, 0, IdSpecs(), headerRow: 1, maxRows: 10, out nint first));
                Assert.Equal(NativeStatus.Ok, TypedApi.OpenTypedReader(live, 0, IdSpecs(), headerRow: 1, maxRows: 7, out nint second));
                try
                {
                    Assert.Equal(NativeStatus.Ok, TypedApi.NextTypedBatch(first, out NativeTable a));
                    Assert.Equal(NativeStatus.Ok, TypedApi.NextTypedBatch(second, out NativeTable b));
                    long fromFirst = a.RowCount;
                    long fromSecond = b.RowCount;
                    TypedApi.FreeTable(ref a);
                    TypedApi.FreeTable(ref b);

                    Assert.Equal(SmallRowCount, fromFirst + Drain(first));
                    Assert.Equal(SmallRowCount, fromSecond + Drain(second));
                }
                finally
                {
                    TypedApi.CloseTypedReader(first);
                    TypedApi.CloseTypedReader(second);
                }
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void ParseTyped_And_An_Arrow_Stream_Should_Leave_A_Live_Reader_Untouched()
        {
            string path = WriteCsv(SmallRowCount);
            try
            {
                Assert.Equal(NativeStatus.Ok, NativeApiTests.OpenPath(path, NativeFormat.Csv, out NativeHandle? handle));
                using NativeHandle live = handle!;
                Assert.Equal(NativeStatus.Ok, TypedApi.OpenTypedReader(live, 0, IdSpecs(), headerRow: 1, maxRows: 10, out nint reader));
                try
                {
                    Assert.Equal(NativeStatus.Ok, TypedApi.NextTypedBatch(reader, out NativeTable first));
                    long taken = first.RowCount;
                    TypedApi.FreeTable(ref first);

                    Assert.Equal(NativeStatus.Ok, TypedApi.ParseTyped(live, 0, IdSpecs(), headerRow: 1, out NativeTable whole));
                    Assert.Equal(SmallRowCount, whole.RowCount);
                    TypedApi.FreeTable(ref whole);
                    Assert.Equal(NativeStatus.Ok, ArrowApi.OpenArrowStream(live, 0, IdSpecs(), headerRow: 1, maxRows: 10, out ArrowArrayStream stream));
                    ReleaseStream(ref stream);

                    Assert.Equal(SmallRowCount, taken + Drain(reader));
                }
                finally
                {
                    TypedApi.CloseTypedReader(reader);
                }
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Theory]
        [InlineData(NativeFormat.Csv)]
        [InlineData(NativeFormat.Xlsx)]
        public void A_Typed_Reader_Opened_But_Not_Yet_Read_Should_Read_Every_Row_After_Its_Workbook_Is_Closed(int format)
        {
            string path = format == NativeFormat.Csv ? WriteCsv(SmallRowCount) : XlsxFixture;
            NativeColumnSpec[] specs = format == NativeFormat.Csv ? IdSpecs() : FirstColumnSpecs();
            int headerRow = format == NativeFormat.Csv ? 1 : 0;
            long expected = format == NativeFormat.Csv ? SmallRowCount : 3;
            try
            {
                Assert.Equal(NativeStatus.Ok, NativeApiTests.OpenPath(path, format, out NativeHandle? handle));
                Assert.Equal(NativeStatus.Ok, TypedApi.OpenTypedReader(handle, 0, specs, headerRow, maxRows: 2, out nint reader));
                try
                {
                    Assert.Equal(NativeStatus.Ok, ReadApi.Close(handle));

                    Assert.Equal(expected, Drain(reader));
                }
                finally
                {
                    TypedApi.CloseTypedReader(reader);
                }
            }
            finally
            {
                if (format == NativeFormat.Csv)
                {
                    File.Delete(path);
                }
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
                    TypedApi.OpenTypedReader(live, 0, IdSpecs(), headerRow: 1, maxRows: 2, out nint reader));
                try
                {
                    Assert.Equal(NativeStatus.Ok, TypedApi.NextTypedBatch(reader, out NativeTable good));
                    Assert.Equal(2L, RowsIn(good));
                    TypedApi.FreeTable(ref good);

                    Assert.Equal(NativeStatus.Error, TypedApi.NextTypedBatch(reader, out NativeTable bad));
                    Assert.Equal(IntPtr.Zero, bad.Columns);
                    string latched = NativeApi.LastErrorText();
                    Assert.NotEmpty(latched);
                    Assert.Contains("failed to convert", latched, StringComparison.Ordinal);

                    Assert.Equal(NativeStatus.Error, TypedApi.NextTypedBatch(reader, out NativeTable again));
                    Assert.Equal(IntPtr.Zero, again.Columns);
                    Assert.Equal(latched, NativeApi.LastErrorText());
                }
                finally
                {
                    TypedApi.CloseTypedReader(reader);
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

                Assert.Equal(NativeStatus.Ok, TypedApi.ParseTyped(live, 0, IdSpecs(), headerRow: 1, out NativeTable first));
                long firstRows = RowsIn(first);
                TypedApi.FreeTable(ref first);
                Assert.Equal(NativeStatus.Ok, TypedApi.ParseTyped(live, 0, IdSpecs(), headerRow: 1, out NativeTable second));
                long secondRows = RowsIn(second);
                TypedApi.FreeTable(ref second);

                Assert.Equal(SmallRowCount, firstRows);
                Assert.Equal(firstRows, secondRows);
            }
            finally
            {
                File.Delete(path);
            }
        }

        private delegate void ReleaseStreamFn(ref ArrowArrayStream stream);

        private static void ReleaseStream(ref ArrowArrayStream stream)
        {
            Marshal.GetDelegateForFunctionPointer<ReleaseStreamFn>(stream.Release)(ref stream);
        }
    }
}
