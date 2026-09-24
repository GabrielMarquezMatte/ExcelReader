using System.Globalization;
using System.Runtime.InteropServices;
using System.Text;
using ExcelReader.Native;
using ExcelReader.Native.Reading;
using ExcelReader.Native.Typed;

namespace ExcelReader.Tests.Native
{
    public sealed partial class NativeApiTests
    {
        [Fact]
        public void ParseTyped_Should_Return_Typed_Columns_By_Name()
        {
            string path = Path.Combine(Path.GetTempPath(), $"excelreader-native-{Guid.NewGuid():N}.csv");
            File.WriteAllText(path, "name,qty,price,active,joined\nwidget,3,9.99,true,2024-01-15\ngadget,7,4.5,false,2024-02-20\n");
            try
            {
                Assert.Equal(NativeStatus.Ok, OpenPath(path, NativeFormat.Csv, out NativeHandle? handle));
                NativeColumnSpec[] specs =
                [
                    new() { Names = ["name"], Type = NativeColumnType.String },
                    new() { Names = ["qty"], Type = NativeColumnType.Int64 },
                    new() { Names = ["price"], Type = NativeColumnType.Float64 },
                    new() { Names = ["active"], Type = NativeColumnType.Bool },
                    new() { Names = ["joined"], Type = NativeColumnType.Date },
                ];
                Assert.Equal(NativeStatus.Ok, TypedApi.ParseTyped(handle, specs, headerRow: 1, out NativeTable table));
                try
                {
                    Assert.Equal(5, table.ColumnCount);
                    Assert.Equal(2, table.RowCount);
                    Assert.Equal(["widget", "gadget"], DecodeStringColumn(ColumnAt(table, 0)));

                    long[] qty = new long[2];
                    Marshal.Copy(ColumnAt(table, 1).Values, qty, 0, 2);
                    Assert.Equal([3L, 7L], qty);

                    double[] prices = new double[2];
                    Marshal.Copy(ColumnAt(table, 2).Values, prices, 0, 2);
                    Assert.Equal([9.99, 4.5], prices);

                    byte[] flags = new byte[2];
                    Marshal.Copy(ColumnAt(table, 3).Values, flags, 0, 2);
                    Assert.Equal([(byte)1, (byte)0], flags);

                    int[] days = new int[2];
                    Marshal.Copy(ColumnAt(table, 4).Values, days, 0, 2);
                    int epoch = new DateOnly(1970, 1, 1).DayNumber;
                    Assert.Equal(new DateOnly(2024, 1, 15).DayNumber - epoch, days[0]);
                    Assert.Equal(new DateOnly(2024, 2, 20).DayNumber - epoch, days[1]);
                }
                finally
                {
                    TypedApi.FreeTable(ref table);
                    ReadApi.Close(handle);
                }
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void ParseTyped_Should_Resolve_By_Index_When_Header_Row_Is_Zero()
        {
            string path = Path.Combine(Path.GetTempPath(), $"excelreader-native-{Guid.NewGuid():N}.csv");
            File.WriteAllText(path, "1,2\n3,4\n");
            try
            {
                Assert.Equal(NativeStatus.Ok, OpenPath(path, NativeFormat.Csv, out NativeHandle? handle));
                NativeColumnSpec[] specs =
                [
                    new() { Index = 0, Type = NativeColumnType.Int64 },
                    new() { Index = 1, Type = NativeColumnType.Int64 },
                ];
                Assert.Equal(NativeStatus.Ok, TypedApi.ParseTyped(handle, specs, headerRow: 0, out NativeTable table));
                try
                {
                    Assert.Equal(2, table.RowCount);
                    long[] first = new long[2];
                    Marshal.Copy(ColumnAt(table, 0).Values, first, 0, 2);
                    long[] second = new long[2];
                    Marshal.Copy(ColumnAt(table, 1).Values, second, 0, 2);
                    Assert.Equal([1L, 3L], first);
                    Assert.Equal([2L, 4L], second);
                }
                finally
                {
                    TypedApi.FreeTable(ref table);
                    ReadApi.Close(handle);
                }
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void ParseTyped_Should_Compute_Time_And_Timestamp_As_Microseconds()
        {
            string path = Path.Combine(Path.GetTempPath(), $"excelreader-native-{Guid.NewGuid():N}.csv");
            File.WriteAllText(path, "at,logged\n13:45:30,2024-01-15T13:45:30\n");
            try
            {
                Assert.Equal(NativeStatus.Ok, OpenPath(path, NativeFormat.Csv, out NativeHandle? handle));
                NativeColumnSpec[] specs =
                [
                    new() { Names = ["at"], Type = NativeColumnType.Time },
                    new() { Names = ["logged"], Type = NativeColumnType.Timestamp },
                ];
                Assert.Equal(NativeStatus.Ok, TypedApi.ParseTyped(handle, specs, headerRow: 1, out NativeTable table));
                try
                {
                    long[] time = new long[1];
                    Marshal.Copy(ColumnAt(table, 0).Values, time, 0, 1);
                    Assert.Equal(new TimeOnly(13, 45, 30).ToTimeSpan().Ticks / 10, time[0]);

                    long[] timestamp = new long[1];
                    Marshal.Copy(ColumnAt(table, 1).Values, timestamp, 0, 1);
                    DateTime expected = new(2024, 1, 15, 13, 45, 30, DateTimeKind.Unspecified);
                    Assert.Equal((expected - DateTime.UnixEpoch).Ticks / 10, timestamp[0]);
                }
                finally
                {
                    TypedApi.FreeTable(ref table);
                    ReadApi.Close(handle);
                }
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Theory]
        [InlineData(int.MinValue, false)]
        [InlineData(-1, false)]
        [InlineData(0, false)]
        [InlineData(1, true)]
        [InlineData(16_384, true)]
        [InlineData(16_385, false)]
        [InlineData(int.MaxValue, false)]
        public void IsValidSpecCount_Should_Accept_Only_One_Through_Excels_Column_Ceiling(int specCount, bool expected)
        {
            Assert.Equal(expected, TypedApi.IsValidSpecCount(specCount));
        }

        [Theory]
        [InlineData(int.MinValue, false)]
        [InlineData(-1, false)]
        [InlineData(0, true)]
        [InlineData(131_068, true)]
        [InlineData(131_069, false)]
        [InlineData(int.MaxValue, false)]
        public void IsValidNameLength_Should_Bound_What_Becomes_A_Read_Length(int nameLength, bool expected)
        {
            Assert.Equal(expected, TypedApi.IsValidNameLength(nameLength));
        }

        [Theory]
        [InlineData(int.MinValue, false)]
        [InlineData(-1, false)]
        [InlineData(0, true)]
        [InlineData(32, true)]
        [InlineData(33, false)]
        [InlineData(int.MaxValue, false)]
        public void IsValidNameCount_Should_Bound_What_Sizes_The_Candidate_Array(int nameCount, bool expected)
        {
            Assert.Equal(expected, TypedApi.IsValidNameCount(nameCount));
        }

        [Fact]
        public void ParseTyped_Should_Reject_A_Blank_Column_Name()
        {
            string path = Path.Combine(Path.GetTempPath(), $"excelreader-native-{Guid.NewGuid():N}.csv");
            File.WriteAllText(path, "name,,qty\nwidget,x,3\n");
            try
            {
                Assert.Equal(NativeStatus.Ok, OpenPath(path, NativeFormat.Csv, out NativeHandle? handle));
                NativeColumnSpec[] specs = [new() { Names = ["   "], Type = NativeColumnType.String }];
                try
                {
                    Assert.Equal(NativeStatus.InvalidArgument, TypedApi.ParseTyped(handle, specs, headerRow: 1, out NativeTable table));
                    Assert.Equal(IntPtr.Zero, table.Columns);

                    Span<byte> buffer = stackalloc byte[256];
                    Assert.Equal(NativeStatus.Ok, NativeApi.LastError(buffer, out int length));
                    Assert.Contains("blank name", Encoding.UTF8.GetString(buffer[..length]), StringComparison.Ordinal);
                }
                finally
                {
                    ReadApi.Close(handle);
                }
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Theory]
        [InlineData(1)]
        [InlineData(7)]
        [InlineData(8)]
        [InlineData(9)]
        [InlineData(17)]
        public void ParseTyped_Validity_Bitmap_Should_Backfill_Rows_Before_The_First_Null(int firstNull)
        {
            const int rowCount = 20;
            StringBuilder csv = new("qty\n");
            for (int i = 0; i < rowCount; i++)
            {
                csv.Append(i == firstNull || i == rowCount - 1 ? "notanumber" : i.ToString(CultureInfo.InvariantCulture)).Append('\n');
            }

            string path = Path.Combine(Path.GetTempPath(), $"excelreader-native-{Guid.NewGuid():N}.csv");
            File.WriteAllText(path, csv.ToString());
            try
            {
                Assert.Equal(NativeStatus.Ok, OpenPath(path, NativeFormat.Csv, out NativeHandle? handle));
                NativeColumnSpec[] specs = [new() { Names = ["qty"], Type = NativeColumnType.Int64, Nullable = true }];
                Assert.Equal(NativeStatus.Ok, TypedApi.ParseTyped(handle, specs, headerRow: 1, out NativeTable table));
                try
                {
                    bool[] expected = new bool[rowCount];
                    Array.Fill(expected, true);
                    expected[firstNull] = false;
                    expected[rowCount - 1] = false;
                    Assert.Equal(expected, DecodeValidity(ColumnAt(table, 0)));
                }
                finally
                {
                    TypedApi.FreeTable(ref table);
                    ReadApi.Close(handle);
                }
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void ParseTyped_Validity_Bitmap_Should_Survive_Byte_Boundaries()
        {
            const int rowCount = 20;
            int[] nullRows = [0, 7, 8, 15, 16, 19];

            StringBuilder csv = new("qty\n");
            for (int i = 0; i < rowCount; i++)
            {
                csv.Append(nullRows.Contains(i) ? "notanumber" : i.ToString(CultureInfo.InvariantCulture)).Append('\n');
            }

            string path = Path.Combine(Path.GetTempPath(), $"excelreader-native-{Guid.NewGuid():N}.csv");
            File.WriteAllText(path, csv.ToString());
            try
            {
                Assert.Equal(NativeStatus.Ok, OpenPath(path, NativeFormat.Csv, out NativeHandle? handle));
                NativeColumnSpec[] specs = [new() { Names = ["qty"], Type = NativeColumnType.Int64, Nullable = true }];
                Assert.Equal(NativeStatus.Ok, TypedApi.ParseTyped(handle, specs, headerRow: 1, out NativeTable table));
                try
                {
                    NativeColumn column = ColumnAt(table, 0);
                    Assert.Equal(rowCount, column.Length);
                    Assert.NotEqual(IntPtr.Zero, column.Validity);

                    bool[] expected = new bool[rowCount];
                    Array.Fill(expected, true);
                    foreach (int row in nullRows)
                    {
                        expected[row] = false;
                    }
                    Assert.Equal(expected, DecodeValidity(column));

                    long[] values = new long[rowCount];
                    Marshal.Copy(column.Values, values, 0, rowCount);
                    for (int i = 0; i < rowCount; i++)
                    {
                        Assert.Equal(nullRows.Contains(i) ? 0L : i, values[i]);
                    }
                }
                finally
                {
                    TypedApi.FreeTable(ref table);
                    ReadApi.Close(handle);
                }
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void ParseTyped_Should_Mark_Failed_Nullable_Conversions_In_The_Validity_Bitmap()
        {
            string path = Path.Combine(Path.GetTempPath(), $"excelreader-native-{Guid.NewGuid():N}.csv");
            File.WriteAllText(path, "qty\n5\n\nnotanumber\n");
            try
            {
                Assert.Equal(NativeStatus.Ok, OpenPath(path, NativeFormat.Csv, out NativeHandle? handle));
                NativeColumnSpec[] specs = [new() { Names = ["qty"], Type = NativeColumnType.Int64, Nullable = true }];
                Assert.Equal(NativeStatus.Ok, TypedApi.ParseTyped(handle, specs, headerRow: 1, out NativeTable table));
                try
                {
                    NativeColumn column = ColumnAt(table, 0);
                    Assert.NotEqual(IntPtr.Zero, column.Validity);
                    Assert.Equal([true, false, false], DecodeValidity(column));

                    long[] values = new long[3];
                    Marshal.Copy(column.Values, values, 0, 3);
                    Assert.Equal(5L, values[0]);
                }
                finally
                {
                    TypedApi.FreeTable(ref table);
                    ReadApi.Close(handle);
                }
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void ParseTyped_Should_Fail_For_A_Non_Nullable_Conversion_Failure()
        {
            string path = Path.Combine(Path.GetTempPath(), $"excelreader-native-{Guid.NewGuid():N}.csv");
            File.WriteAllText(path, "qty\n5\nnotanumber\n");
            try
            {
                Assert.Equal(NativeStatus.Ok, OpenPath(path, NativeFormat.Csv, out NativeHandle? handle));
                NativeColumnSpec[] specs = [new() { Names = ["qty"], Type = NativeColumnType.Int64, Nullable = false }];

                int status = TypedApi.ParseTyped(handle, specs, headerRow: 1, out NativeTable table);

                Assert.Equal(NativeStatus.Error, status);
                Assert.Equal(IntPtr.Zero, table.Columns);
                Span<byte> buffer = stackalloc byte[256];
                Assert.Equal(NativeStatus.Ok, NativeApi.LastError(buffer, out int length));
                Assert.True(length > 0);
                ReadApi.Close(handle);
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void ParseTyped_Should_Reject_A_Name_Based_Spec_When_Header_Row_Is_Zero()
        {
            Assert.Equal(NativeStatus.Ok, OpenPath(XlsxFixture, NativeFormat.Auto, out NativeHandle? handle));
            try
            {
                NativeColumnSpec[] specs = [new() { Names = ["anything"], Type = NativeColumnType.String }];
                Assert.Equal(NativeStatus.InvalidArgument, TypedApi.ParseTyped(handle, specs, headerRow: 0, out NativeTable table));
                Assert.Equal(IntPtr.Zero, table.Columns);
            }
            finally
            {
                ReadApi.Close(handle);
            }
        }

        [Fact]
        public void ParseTyped_Should_Reject_An_Unmatched_Header_Name()
        {
            string path = Path.Combine(Path.GetTempPath(), $"excelreader-native-{Guid.NewGuid():N}.csv");
            File.WriteAllText(path, "name,qty\nwidget,3\n");
            try
            {
                Assert.Equal(NativeStatus.Ok, OpenPath(path, NativeFormat.Csv, out NativeHandle? handle));
                NativeColumnSpec[] specs = [new() { Names = ["does-not-exist"], Type = NativeColumnType.String }];

                int status = TypedApi.ParseTyped(handle, specs, headerRow: 1, out NativeTable table);

                Assert.Equal(NativeStatus.InvalidArgument, status);
                Assert.Equal(IntPtr.Zero, table.Columns);
                ReadApi.Close(handle);
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void ParseTyped_Should_Resolve_The_First_Candidate_Name_Present_In_The_Header()
        {
            string path = Path.Combine(Path.GetTempPath(), $"excelreader-native-{Guid.NewGuid():N}.csv");
            File.WriteAllText(path, "qty,quantity\n0,5\n");
            try
            {
                Assert.Equal(NativeStatus.Ok, OpenPath(path, NativeFormat.Csv, out NativeHandle? handle));
                NativeColumnSpec[] specs = [new() { Names = ["does-not-exist", "quantity"], Type = NativeColumnType.Int64 }];
                int status = TypedApi.ParseTyped(handle, specs, headerRow: 1, out NativeTable table);
                Assert.Equal(NativeStatus.Ok, status);
                Assert.Equal(1, table.RowCount);
                long value = Marshal.ReadInt64(ColumnAt(table, 0).Values);
                Assert.Equal(5, value);
                TypedApi.FreeTable(ref table);
                ReadApi.Close(handle);
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void ParseTyped_Should_Fail_With_A_Message_Listing_Every_Candidate_When_None_Match()
        {
            string path = Path.Combine(Path.GetTempPath(), $"excelreader-native-{Guid.NewGuid():N}.csv");
            File.WriteAllText(path, "qty\n5\n");
            try
            {
                Assert.Equal(NativeStatus.Ok, OpenPath(path, NativeFormat.Csv, out NativeHandle? handle));
                NativeColumnSpec[] specs = [new() { Names = ["nope", "still-nope"], Type = NativeColumnType.Int64 }];

                int status = TypedApi.ParseTyped(handle, specs, headerRow: 1, out NativeTable table);

                Assert.Equal(NativeStatus.InvalidArgument, status);
                Assert.Equal(IntPtr.Zero, table.Columns);
                Span<byte> buffer = stackalloc byte[256];
                NativeApi.LastError(buffer, out int length);
                string message = Encoding.UTF8.GetString(buffer[..length]);
                Assert.Contains("\"nope\"", message, StringComparison.Ordinal);
                Assert.Contains("\"still-nope\"", message, StringComparison.Ordinal);
                ReadApi.Close(handle);
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void ParseTyped_Should_Reject_Zero_Specs()
        {
            Assert.Equal(NativeStatus.Ok, OpenPath(XlsxFixture, NativeFormat.Auto, out NativeHandle? handle));
            try
            {
                Assert.Equal(NativeStatus.InvalidArgument, TypedApi.ParseTyped(handle, [], headerRow: 1, out NativeTable table));
                Assert.Equal(IntPtr.Zero, table.Columns);
            }
            finally
            {
                ReadApi.Close(handle);
            }
        }

        [Fact]
        public void ParseTyped_Should_Reject_A_Null_Handle()
        {
            NativeColumnSpec[] specs = [new() { Index = 0, Type = NativeColumnType.String }];
            Assert.Equal(NativeStatus.InvalidHandle, TypedApi.ParseTyped(null, specs, headerRow: 1, out _));
        }

        [Fact]
        public void ParseTyped_Should_Not_Disturb_The_Row_Cursor()
        {
            string path = Path.Combine(Path.GetTempPath(), $"excelreader-native-{Guid.NewGuid():N}.csv");
            File.WriteAllText(path, "name\nfirst\nsecond\n");
            try
            {
                Assert.Equal(NativeStatus.Ok, OpenPath(path, NativeFormat.Csv, out NativeHandle? handle));
                try
                {
                    Assert.Equal(NativeStatus.Ok, ReadApi.NextRow(handle, new byte[4096], out _));

                    NativeColumnSpec[] specs = [new() { Names = ["name"], Type = NativeColumnType.String }];
                    Assert.Equal(NativeStatus.Ok, TypedApi.ParseTyped(handle, specs, headerRow: 1, out NativeTable table));
                    TypedApi.FreeTable(ref table);

                    byte[] buffer = new byte[4096];
                    Assert.Equal(NativeStatus.Ok, ReadApi.NextRow(handle, buffer, out int written));
                    Assert.Equal("first", DecodeRow(buffer.AsSpan(0, written))[0].Value);
                }
                finally
                {
                    ReadApi.Close(handle);
                }
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void FreeTable_Should_Be_Idempotent_On_A_Zeroed_Table()
        {
            NativeTable table = default;
            TypedApi.FreeTable(ref table);
            TypedApi.FreeTable(ref table);
            Assert.Equal(IntPtr.Zero, table.Columns);
        }
    }
}
