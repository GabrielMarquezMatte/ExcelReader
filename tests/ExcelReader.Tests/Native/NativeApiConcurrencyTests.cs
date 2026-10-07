using System.Runtime.InteropServices;
using System.Text;
using ExcelReader.Native;
using ExcelReader.Native.Reading;
using ExcelReader.Native.Typed;
using ExcelReader.Tests.Reader;

namespace ExcelReader.Tests.Native
{
    public class NativeApiConcurrencyTests
    {
        private static readonly string XlsxFixture = Path.Combine(AppContext.BaseDirectory, "data", "sample.xlsx");

        [Fact]
        public Task IndependentHandlesOnSeparateThreadsDoNotInterfere()
        {
            const int handleCount = 16;
            IEnumerable<Task> tasks = Enumerable.Range(0, handleCount).Select(_ => Task.Run(() =>
            {
                Assert.Equal(NativeStatus.Ok, ReadApi.OpenFile(Encoding.UTF8.GetBytes(XlsxFixture), NativeFormat.Auto, out NativeHandle? handle));
                Assert.NotNull(handle);
                using NativeRowCursor cursor = NativeApiTests.OpenRows(handle);

                int rowCount = 0;
                while (ReadApi.NextRow(cursor, new byte[64 * 1024], out _) != NativeStatus.Eof)
                {
                    rowCount++;
                }

                Assert.Equal(3, rowCount);
                Assert.Equal(NativeStatus.Ok, ReadApi.Close(handle));
            }));

            return Task.WhenAll(tasks);
        }

        [Fact]
        public Task LastErrorIsPerThreadNotSharedAcrossHandles()
        {
            const int threadCount = 16;
            IEnumerable<Task> tasks = Enumerable.Range(0, threadCount).Select(i => Task.Run(() =>
            {
                if (i % 2 == 0)
                {
                    byte[] missingPath = Encoding.UTF8.GetBytes($"does-not-exist-{i}.xlsx");
                    int status = ReadApi.OpenFile(missingPath, NativeFormat.Auto, out NativeHandle? handle);
                    Assert.NotEqual(NativeStatus.Ok, status);
                    Assert.Null(handle);

                    Span<byte> errorBuffer = stackalloc byte[256];
                    Assert.Equal(NativeStatus.Ok, NativeApi.LastError(errorBuffer, out int length));
                    Assert.True(length > 0);
                }
                else
                {
                    Assert.Equal(NativeStatus.Ok, ReadApi.OpenFile(Encoding.UTF8.GetBytes(XlsxFixture), NativeFormat.Auto, out NativeHandle? handle));
                    Assert.NotNull(handle);

                    Span<byte> errorBuffer = stackalloc byte[256];
                    Assert.Equal(NativeStatus.Ok, NativeApi.LastError(errorBuffer, out int length));
                    Assert.Equal(0, length);

                    Assert.Equal(NativeStatus.Ok, ReadApi.Close(handle));
                }
            }));

            return Task.WhenAll(tasks);
        }

        private static List<string> ReadSheet(NativeHandle handle, int sheet)
        {
            Assert.Equal(NativeStatus.Ok, ReadApi.OpenRows(handle, sheet, out NativeRowCursor? opened));
            using NativeRowCursor cursor = opened!;
            List<string> lines = [];
            while (true)
            {
                int status = ReadApi.NextRowView(cursor, out NativeRow row);
                if (status == NativeStatus.Eof)
                {
                    return lines;
                }
                Assert.Equal(NativeStatus.Ok, status);
                StringBuilder line = new();
                int cellSize = Marshal.SizeOf<NativeRowCell>();
                for (int i = 0; i < row.CellCount; i++)
                {
                    NativeRowCell cell = Marshal.PtrToStructure<NativeRowCell>(IntPtr.Add(row.Cells, i * cellSize));
                    line.Append(Marshal.PtrToStringUTF8(cell.Value, cell.ValueLength)).Append('|');
                }
                lines.Add(line.ToString());
            }
        }

        [Theory]
        [InlineData(ConcurrentSheetTests.Format.Xlsx, NativeFormat.Xlsx)]
        [InlineData(ConcurrentSheetTests.Format.Xlsb, NativeFormat.Xlsb)]
        [InlineData(ConcurrentSheetTests.Format.Xls, NativeFormat.Xls)]
        public void CursorsOverEverySheetOfOneWorkbookReadTheSameRowsInParallelAsInSequence(ConcurrentSheetTests.Format format, int nativeFormat)
        {
            byte[] bytes = ConcurrentSheetTests.Build(format);
            Assert.Equal(NativeStatus.Ok, ReadApi.OpenMemory(bytes, nativeFormat, out NativeHandle? handle));
            using NativeHandle live = handle!;
            List<string>[] sequential = new List<string>[ConcurrentSheetTests.Sheets];
            for (int s = 0; s < sequential.Length; s++)
            {
                sequential[s] = ReadSheet(live, s);
                Assert.Equal(ConcurrentSheetTests.Rows, sequential[s].Count);
            }

            List<string>[] parallel = new List<string>[ConcurrentSheetTests.Sheets * 2];
            Parallel.For(0, parallel.Length, i => parallel[i] = ReadSheet(live, i % ConcurrentSheetTests.Sheets));

            for (int i = 0; i < parallel.Length; i++)
            {
                Assert.Equal(sequential[i % ConcurrentSheetTests.Sheets], parallel[i]);
            }
        }

        [Fact]
        public void TypedParsesOfEverySheetOfOneWorkbookRunInParallel()
        {
            byte[] bytes = ConcurrentSheetTests.Build(ConcurrentSheetTests.Format.Xlsx);
            Assert.Equal(NativeStatus.Ok, ReadApi.OpenMemory(bytes, NativeFormat.Xlsx, out NativeHandle? handle));
            using NativeHandle live = handle!;
            long[] rowCounts = new long[ConcurrentSheetTests.Sheets];
            long[] firstValues = new long[ConcurrentSheetTests.Sheets];

            Parallel.For(0, rowCounts.Length, s =>
            {
                NativeColumnSpec[] specs = [new() { Index = 1, Type = NativeColumnType.Int64 }];
                Assert.Equal(NativeStatus.Ok, TypedApi.ParseTyped(live, s, specs, headerRow: 0, out NativeTable table));
                try
                {
                    rowCounts[s] = table.RowCount;
                    NativeColumn column = Marshal.PtrToStructure<NativeColumn>(table.Columns);
                    firstValues[s] = Marshal.ReadInt64(column.Values);
                }
                finally
                {
                    TypedApi.FreeTable(ref table);
                }
            });

            Assert.All(rowCounts, count => Assert.Equal(ConcurrentSheetTests.Rows, count));
            Assert.Equal(Enumerable.Range(0, firstValues.Length).Select(s => s * 1_000_000L), firstValues);
        }

        [Fact]
        public void AClosedWorkbookIdStaysInvalidWhileItsCursorReadsToTheEnd()
        {
            byte[] bytes = ConcurrentSheetTests.Build(ConcurrentSheetTests.Format.Xlsb);
            Assert.Equal(NativeStatus.Ok, ReadApi.OpenMemory(bytes, NativeFormat.Xlsb, out NativeHandle? handle));
            nint id = NativeHandleTable.Register(handle!);
            Assert.Equal(NativeStatus.Ok, ReadApi.OpenRows(Exports.Resolve(id), 3, out NativeRowCursor? opened));
            using NativeRowCursor cursor = opened!;

            Assert.True(Exports.TryFree(id, out NativeHandle? freed));
            Assert.Equal(NativeStatus.Ok, ReadApi.Close(freed));

            Assert.Null(Exports.Resolve(id));
            Assert.Equal(NativeStatus.InvalidHandle, ReadApi.OpenRows(Exports.Resolve(id), 0, out _));
            int rows = 0;
            while (ReadApi.NextRowView(cursor, out NativeRow row) == NativeStatus.Ok)
            {
                if (rows == 0)
                {
                    NativeRowCell first = Marshal.PtrToStructure<NativeRowCell>(row.Cells);
                    Assert.StartsWith("s3-name-", Marshal.PtrToStringUTF8(first.Value, first.ValueLength), StringComparison.Ordinal);
                }
                rows++;
            }
            Assert.Equal(ConcurrentSheetTests.Rows, rows);
        }
    }
}
