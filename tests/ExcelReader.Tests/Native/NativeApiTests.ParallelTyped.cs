using System.Globalization;
using System.Runtime.InteropServices;
using System.Text;
using ExcelReader.Native;
using ExcelReader.Native.Arrow;
using ExcelReader.Native.Reading;
using ExcelReader.Native.Typed;

namespace ExcelReader.Tests.Native
{
    public sealed partial class NativeApiTests
    {
        private const int SmallChunk = 64 * 1024;
        private const int MixedRows = 20_000;

        private static readonly NativeColumnSpec[] MixedSpecs =
        [
            new() { Names = ["name"], Type = NativeColumnType.String },
            new() { Names = ["qty"], Type = NativeColumnType.Int64, Nullable = true },
            new() { Names = ["price"], Type = NativeColumnType.Float64 },
            new() { Names = ["flag"], Type = NativeColumnType.Bool },
            new() { Names = ["day"], Type = NativeColumnType.Date },
            new() { Names = ["note"], Type = NativeColumnType.String },
        ];

        private static string MixedCsv(int rows, string newline = "\n", bool bom = false, bool trailingNewline = true, char delimiter = ',')
        {
            StringBuilder csv = new(bom ? "﻿" : "");
            csv.Append(string.Join(delimiter, "name", "qty", "price", "flag", "day", "note")).Append(newline);
            for (int i = 0; i < rows; i++)
            {
                string qty = i % 7 == 0 ? "" : (i * 3).ToString(CultureInfo.InvariantCulture);
                string note = i % 50 == 0
                    ? string.Create(CultureInfo.InvariantCulture, $"\"multi{newline}line{delimiter} {i}\"")
                    : string.Create(CultureInfo.InvariantCulture, $"n{i}");
                csv.Append(CultureInfo.InvariantCulture, $"item{i}{delimiter}{qty}{delimiter}{i * 0.25}{delimiter}{(i % 2 == 0 ? "true" : "false")}{delimiter}2024-01-{1 + (i % 28):D2}{delimiter}{note}");
                if (i < rows - 1 || trailingNewline)
                {
                    csv.Append(newline);
                }
            }
            return csv.ToString();
        }

        private static (int Status, NativeTable Table, string Error, bool Parallel) ParseWith(
            byte[] csv, NativeColumnSpec[] specs, int headerRow, int dop, int chunkSize, NativeOpenOptions? options = null)
        {
            Assert.Equal(NativeStatus.Ok, ReadApi.OpenMemory(csv, NativeFormat.Csv, out NativeHandle? handle, options));
            try
            {
                int status = TypedApi.ParseTypedTable(handle, specs, headerRow, dop, "test", out NativeTable table, chunkSize);
                return (status, table, NativeApi.LastErrorText(), TypedApi.LastParseRanInParallel);
            }
            finally
            {
                ReadApi.Close(handle);
            }
        }

        private static byte[] ReadBlock(IntPtr pointer, int length)
        {
            byte[] bytes = new byte[length];
            Marshal.Copy(pointer, bytes, 0, length);
            return bytes;
        }

        private static void AssertSameTable(NativeTable expected, NativeTable actual)
        {
            Assert.Equal(expected.ColumnCount, actual.ColumnCount);
            Assert.Equal(expected.RowCount, actual.RowCount);
            for (int c = 0; c < expected.ColumnCount; c++)
            {
                NativeColumn e = ColumnAt(expected, c);
                NativeColumn a = ColumnAt(actual, c);
                Assert.Equal(e.Type, a.Type);
                Assert.Equal(e.Length, a.Length);
                Assert.Equal(e.Validity == IntPtr.Zero, a.Validity == IntPtr.Zero);
                int bitmapBytes = (int)((e.Length + 7) / 8);
                if (e.Validity != IntPtr.Zero)
                {
                    Assert.Equal(ReadBlock(e.Validity, bitmapBytes), ReadBlock(a.Validity, bitmapBytes));
                }
                int valueBytes = e.Type switch
                {
                    NativeColumnType.String => (int)(e.Length + 1) * sizeof(int),
                    NativeColumnType.Bool => (int)e.Length,
                    NativeColumnType.Date => (int)e.Length * sizeof(int),
                    _ => (int)e.Length * sizeof(long),
                };
                Assert.Equal(ReadBlock(e.Values, valueBytes), ReadBlock(a.Values, valueBytes));
                Assert.Equal(e.DataLen, a.DataLen);
                if (e.DataLen > 0)
                {
                    Assert.Equal(ReadBlock(e.Data, (int)e.DataLen), ReadBlock(a.Data, (int)a.DataLen));
                }
            }
        }

        private static void AssertParallelMatchesSequential(byte[] csv, NativeColumnSpec[] specs, int headerRow, long expectedRows, NativeOpenOptions? options = null)
        {
            (int seqStatus, NativeTable sequential, _, _) = ParseWith(csv, specs, headerRow, dop: 1, chunkSize: 0, options);
            (int parStatus, NativeTable parallel, _, bool ranInParallel) = ParseWith(csv, specs, headerRow, dop: 4, SmallChunk, options);
            try
            {
                Assert.Equal(NativeStatus.Ok, seqStatus);
                Assert.Equal(NativeStatus.Ok, parStatus);
                Assert.True(ranInParallel);
                Assert.Equal(expectedRows, sequential.RowCount);
                AssertSameTable(sequential, parallel);
            }
            finally
            {
                TypedApi.FreeTable(ref sequential);
                TypedApi.FreeTable(ref parallel);
            }
        }

        [Theory]
        [InlineData("\n", false, true)]
        [InlineData("\r\n", true, true)]
        [InlineData("\n", false, false)]
        public void ParseTypedTable_Should_Match_The_Sequential_Table_Across_Partitions(string newline, bool bom, bool trailingNewline)
        {
            byte[] csv = Encoding.UTF8.GetBytes(MixedCsv(MixedRows, newline, bom, trailingNewline));
            AssertParallelMatchesSequential(csv, MixedSpecs, headerRow: 1, MixedRows);
        }

        [Fact]
        public void ParseTypedTable_Should_Merge_Validity_From_Partitions_With_And_Without_Nulls()
        {
            StringBuilder text = new("head,tail,none\n");
            for (int i = 0; i < 60_000; i++)
            {
                string head = i < 777 && i % 3 == 0 ? "" : i.ToString(CultureInfo.InvariantCulture);
                string tail = i > 59_000 && i % 5 == 0 ? "" : i.ToString(CultureInfo.InvariantCulture);
                text.Append(CultureInfo.InvariantCulture, $"{head},{tail},{i}\n");
            }
            NativeColumnSpec[] specs =
            [
                new() { Names = ["head"], Type = NativeColumnType.Int64, Nullable = true },
                new() { Names = ["tail"], Type = NativeColumnType.Int64, Nullable = true },
                new() { Names = ["none"], Type = NativeColumnType.Int64, Nullable = true },
            ];
            AssertParallelMatchesSequential(Encoding.UTF8.GetBytes(text.ToString()), specs, headerRow: 1, expectedRows: 60_000);
        }

        [Theory]
        [InlineData(1)]
        [InlineData(2)]
        [InlineData(3)]
        [InlineData(4)]
        [InlineData(5)]
        public void ParseTypedTable_Should_Free_Every_Block_Exactly_Once_When_A_Parallel_Build_Allocation_Fails(int failAt)
        {
            StringBuilder text = new("qty,name\n");
            for (int i = 0; i < 40_000; i++)
            {
                text.Append(i % 10 == 0 ? "" : i.ToString(CultureInfo.InvariantCulture)).Append(",n").Append(i).Append('\n');
            }
            NativeColumnSpec[] specs =
            [
                new() { Names = ["qty"], Type = NativeColumnType.Int64, Nullable = true },
                new() { Names = ["name"], Type = NativeColumnType.String },
            ];
            BlockTracker tracker = new(failAt);
            Assert.Equal(NativeStatus.Ok, ReadApi.OpenMemory(Encoding.UTF8.GetBytes(text.ToString()), NativeFormat.Csv, out NativeHandle? handle));
            int status;
            NativeTable table;
            TypedApi.AllocOverride = tracker.Alloc;
            TypedApi.FreeOverride = tracker.Free;
            try
            {
                status = TypedApi.ParseTypedTable(handle, specs, headerRow: 1, 4, "test", out table, SmallChunk);
            }
            finally
            {
                TypedApi.AllocOverride = null;
                TypedApi.FreeOverride = null;
                ReadApi.Close(handle);
            }

            Assert.True(TypedApi.LastParseRanInParallel);
            Assert.Equal(NativeStatus.Error, status);
            Assert.Equal(IntPtr.Zero, table.Columns);
            Assert.Equal(failAt, tracker.Allocations);
            Assert.Equal(failAt - 1, tracker.Frees);
            Assert.Equal(0, tracker.Live);
            Assert.Equal(0, tracker.UnknownFrees);
        }

        [Fact]
        public void ParseTypedTable_Should_Use_The_Dialect_The_Handle_Was_Opened_With()
        {
            byte[] csv = Encoding.UTF8.GetBytes(MixedCsv(MixedRows, delimiter: ';'));
            AssertParallelMatchesSequential(csv, MixedSpecs, headerRow: 1, MixedRows, new NativeOpenOptions { CsvDelimiter = (byte)';' });
        }

        [Fact]
        public void ParseTypedTable_Should_Keep_The_Bom_When_The_Handle_Does_Not_Detect_It()
        {
            byte[] csv = Encoding.UTF8.GetBytes(MixedCsv(MixedRows, bom: true));
            NativeColumnSpec[] specs = [new() { Index = 0, Type = NativeColumnType.String }];
            AssertParallelMatchesSequential(csv, specs, headerRow: 0, MixedRows + 1, new NativeOpenOptions { CsvDetectByteOrderMark = false });
        }

        [Fact]
        public void ParseTypedTable_Should_Honor_A_Header_Row_Below_The_First_Line()
        {
            byte[] csv = Encoding.UTF8.GetBytes("exported 2026-09-25\n" + MixedCsv(MixedRows));
            AssertParallelMatchesSequential(csv, MixedSpecs, headerRow: 2, MixedRows);
        }

        [Fact]
        public void ParseTypedTable_Should_Ignore_A_Failure_Seen_Only_From_A_Wrong_Partition_Start()
        {
            StringBuilder text = new("id,note\n");
            for (int record = 1; record <= 40; record++)
            {
                text.Append(CultureInfo.InvariantCulture, $"{record},\"");
                for (int line = 0; line < 2000; line++)
                {
                    text.Append("x,not-a-number\n");
                }
                text.Append("\"\n");
            }
            NativeColumnSpec[] specs =
            [
                new() { Names = ["id"], Type = NativeColumnType.Int64 },
                new() { Names = ["note"], Type = NativeColumnType.String },
            ];
            AssertParallelMatchesSequential(Encoding.UTF8.GetBytes(text.ToString()), specs, headerRow: 1, expectedRows: 40);
        }

        [Fact]
        public void ParseTypedTable_Should_Report_The_First_Failure_In_File_Order()
        {
            StringBuilder text = new("a,b\n");
            for (int i = 0; i < 60_000; i++)
            {
                string a = i == 45_000 ? "bad-a" : i.ToString(CultureInfo.InvariantCulture);
                string b = i == 30_000 ? "bad-b" : i.ToString(CultureInfo.InvariantCulture);
                text.Append(a).Append(',').Append(b).Append('\n');
            }
            byte[] csv = Encoding.UTF8.GetBytes(text.ToString());
            NativeColumnSpec[] specs =
            [
                new() { Names = ["a"], Type = NativeColumnType.Int64 },
                new() { Names = ["b"], Type = NativeColumnType.Int64 },
            ];

            (int seqStatus, _, string seqError, _) = ParseWith(csv, specs, headerRow: 1, dop: 1, chunkSize: 0);
            (int parStatus, NativeTable parallel, string parError, bool ranInParallel) = ParseWith(csv, specs, headerRow: 1, dop: 4, SmallChunk);

            Assert.Equal(NativeStatus.Error, seqStatus);
            Assert.Equal(seqStatus, parStatus);
            Assert.True(ranInParallel);
            Assert.Contains("column 1", seqError, StringComparison.Ordinal);
            Assert.Equal(seqError, parError);
            Assert.Equal(IntPtr.Zero, parallel.Columns);
        }

        [Fact]
        public void ParseTypedTable_Should_Read_A_Small_Csv_Sequentially()
        {
            (int status, NativeTable table, _, bool ranInParallel) = ParseWith(
                Encoding.UTF8.GetBytes(MixedCsv(100)), MixedSpecs, headerRow: 1, dop: 0, chunkSize: 0);
            try
            {
                Assert.Equal(NativeStatus.Ok, status);
                Assert.False(ranInParallel);
                Assert.Equal(100, table.RowCount);
            }
            finally
            {
                TypedApi.FreeTable(ref table);
            }
        }

        [Fact]
        public void ParseTypedTable_Should_Read_A_Non_Csv_Workbook_Like_ParseTyped()
        {
            NativeColumnSpec[] specs = [new() { Index = 0, Type = NativeColumnType.String, Nullable = true }];
            Assert.Equal(NativeStatus.Ok, OpenPath(XlsxFixture, NativeFormat.Auto, out NativeHandle? handle));
            try
            {
                Assert.Equal(NativeStatus.Ok, TypedApi.ParseTyped(handle, specs, headerRow: 0, out NativeTable expected));
                Assert.Equal(NativeStatus.Ok, TypedApi.ParseTypedTable(handle, specs, headerRow: 0, 0, "test", out NativeTable actual));
                try
                {
                    Assert.False(TypedApi.LastParseRanInParallel);
                    AssertSameTable(expected, actual);
                }
                finally
                {
                    TypedApi.FreeTable(ref expected);
                    TypedApi.FreeTable(ref actual);
                }
            }
            finally
            {
                ReadApi.Close(handle);
            }
        }

        [Fact]
        public void ParseTypedTable_Should_Reject_A_Negative_Degree_Of_Parallelism()
        {
            (int status, NativeTable table, string error, _) = ParseWith(
                Encoding.UTF8.GetBytes(MixedCsv(10)), MixedSpecs, headerRow: 1, dop: -1, chunkSize: 0);

            Assert.Equal(NativeStatus.InvalidArgument, status);
            Assert.Contains("degree_of_parallelism", error, StringComparison.Ordinal);
            Assert.Equal(IntPtr.Zero, table.Columns);
        }

        [Fact]
        public void ParseTypedTable_Should_Read_From_A_File_Without_Moving_The_Row_Cursor()
        {
            string path = Path.Combine(Path.GetTempPath(), $"excelreader-native-{Guid.NewGuid():N}.csv");
            byte[] csv = Encoding.UTF8.GetBytes(MixedCsv(MixedRows));
            File.WriteAllBytes(path, csv);
            (_, NativeTable expected, _, _) = ParseWith(csv, MixedSpecs, headerRow: 1, dop: 1, chunkSize: 0);
            Assert.Equal(NativeStatus.Ok, OpenPath(path, NativeFormat.Csv, out NativeHandle? handle));
            try
            {
                byte[] blob = new byte[1 << 16];
                const int advanced = 5_000;
                for (int i = 0; i < advanced; i++)
                {
                    Assert.Equal(NativeStatus.Ok, ReadApi.NextRow(handle, blob, out _));
                }

                Assert.Equal(NativeStatus.Ok, TypedApi.ParseTypedTable(handle, MixedSpecs, headerRow: 1, 4, "test", out NativeTable actual, SmallChunk));
                try
                {
                    Assert.True(TypedApi.LastParseRanInParallel);
                    AssertSameTable(expected, actual);
                }
                finally
                {
                    TypedApi.FreeTable(ref actual);
                }

                Assert.Equal(NativeStatus.Ok, ReadApi.NextRow(handle, blob, out int written));
                Assert.Equal("item4999", DecodeRow(blob.AsSpan(0, written))[0].Value);
                int remaining = 1;
                string last = "";
                while (ReadApi.NextRow(handle, blob, out written) == NativeStatus.Ok)
                {
                    remaining++;
                    last = DecodeRow(blob.AsSpan(0, written))[0].Value;
                }
                Assert.Equal(MixedRows + 1 - advanced, remaining);
                Assert.Equal("item19999", last);
            }
            finally
            {
                TypedApi.FreeTable(ref expected);
                ReadApi.Close(handle);
                File.Delete(path);
            }
        }

        [Fact]
        public void ParseTypedTable_Should_Fault_A_Live_Chunked_Read_On_The_Same_Handle()
        {
            byte[] csv = Encoding.UTF8.GetBytes(MixedCsv(MixedRows));
            Assert.Equal(NativeStatus.Ok, ReadApi.OpenMemory(csv, NativeFormat.Csv, out NativeHandle? handle));
            try
            {
                Assert.Equal(NativeStatus.Ok, TypedApi.OpenTypedReader(handle, MixedSpecs, headerRow: 1, 100, out nint reader));
                try
                {
                    Assert.Equal(NativeStatus.Ok, TypedApi.ParseTypedTable(handle, MixedSpecs, headerRow: 1, 4, "test", out NativeTable whole, SmallChunk));
                    TypedApi.FreeTable(ref whole);

                    Assert.Equal(NativeStatus.Error, TypedApi.NextTypedBatch(reader, out NativeTable after));
                    Assert.Equal(IntPtr.Zero, after.Columns);
                }
                finally
                {
                    TypedApi.CloseTypedReader(reader);
                }
            }
            finally
            {
                ReadApi.Close(handle);
            }
        }

        [Fact]
        public void ParseArrow_Should_Accept_A_Degree_Of_Parallelism()
        {
            byte[] csv = Encoding.UTF8.GetBytes(MixedCsv(MixedRows));
            Assert.Equal(NativeStatus.Ok, ReadApi.OpenMemory(csv, NativeFormat.Csv, out NativeHandle? handle));
            try
            {
                Assert.Equal(NativeStatus.Ok, ArrowApi.ParseArrow(handle, MixedSpecs, headerRow: 1, 0, out ArrowArray array, out ArrowSchema schema));
                try
                {
                    Assert.Equal(MixedRows, array.Length);
                    Assert.Equal(MixedSpecs.Length, array.NChildren);
                }
                finally
                {
                    ExercisedReleaseArrow(ref array, ref schema);
                }
            }
            finally
            {
                ReadApi.Close(handle);
            }
        }
    }
}
