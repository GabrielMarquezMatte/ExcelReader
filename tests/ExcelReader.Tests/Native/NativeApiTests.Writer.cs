using System.IO.Compression;
using System.Runtime.InteropServices;
using System.Text;
using ExcelReader.Core.Reader;
using ExcelReader.Core.Reader.Xlsb;
using ExcelReader.Core.Reader.Xlsx;
using ExcelReader.Core.Writer.Csv;
using ExcelReader.Native;
using ExcelReader.Native.Reading;
using ExcelReader.Native.Typed;
using ExcelReader.Native.Writer;

namespace ExcelReader.Tests.Native
{
    public sealed partial class NativeApiTests
    {
        private static NativeTable BuildInt64Table(long[] values, byte[]? validity = null)
        {
            NativeColumn column = new()
            {
                Type = NativeColumnType.Int64,
                Length = values.LongLength,
                Values = Marshal.AllocHGlobal(values.Length * sizeof(long)),
                Validity = IntPtr.Zero,
                Data = IntPtr.Zero,
                DataLen = 0,
            };
            Marshal.Copy(values, 0, column.Values, values.Length);
            if (validity is not null)
            {
                column.Validity = Marshal.AllocHGlobal(validity.Length);
                Marshal.Copy(validity, 0, column.Validity, validity.Length);
            }

            IntPtr columns = Marshal.AllocHGlobal(Marshal.SizeOf<NativeColumn>());
            Marshal.StructureToPtr(column, columns, false);
            return new NativeTable { ColumnCount = 1, RowCount = values.LongLength, Columns = columns };
        }

        private static NativeTable BuildStringTable(int[] offsets, byte[] data)
        {
            NativeColumn column = new()
            {
                Type = NativeColumnType.String,
                Length = offsets.Length - 1,
                Values = Marshal.AllocHGlobal(offsets.Length * sizeof(int)),
                Validity = IntPtr.Zero,
                Data = data.Length == 0 ? IntPtr.Zero : Marshal.AllocHGlobal(data.Length),
                DataLen = data.Length,
            };
            Marshal.Copy(offsets, 0, column.Values, offsets.Length);
            if (data.Length > 0)
            {
                Marshal.Copy(data, 0, column.Data, data.Length);
            }

            IntPtr columns = Marshal.AllocHGlobal(Marshal.SizeOf<NativeColumn>());
            Marshal.StructureToPtr(column, columns, false);
            return new NativeTable { ColumnCount = 1, RowCount = offsets.Length - 1, Columns = columns };
        }

        private static NativeTable BuildQtyNameTable(long[] quantities, string[] names)
        {
            byte[] data = Encoding.UTF8.GetBytes(string.Concat(names));
            int[] offsets = new int[names.Length + 1];
            for (int index = 0; index < names.Length; index++)
            {
                offsets[index + 1] = offsets[index] + Encoding.UTF8.GetByteCount(names[index]);
            }

            NativeColumn qty = new()
            {
                Type = NativeColumnType.Int64,
                Length = quantities.LongLength,
                Values = Marshal.AllocHGlobal(quantities.Length * sizeof(long)),
            };
            Marshal.Copy(quantities, 0, qty.Values, quantities.Length);
            NativeColumn name = new()
            {
                Type = NativeColumnType.String,
                Length = names.LongLength,
                Values = Marshal.AllocHGlobal(offsets.Length * sizeof(int)),
                Data = data.Length == 0 ? IntPtr.Zero : Marshal.AllocHGlobal(data.Length),
                DataLen = data.Length,
            };
            Marshal.Copy(offsets, 0, name.Values, offsets.Length);
            if (data.Length > 0)
            {
                Marshal.Copy(data, 0, name.Data, data.Length);
            }

            int size = Marshal.SizeOf<NativeColumn>();
            IntPtr block = Marshal.AllocHGlobal(size * 2);
            Marshal.StructureToPtr(qty, block, false);
            Marshal.StructureToPtr(name, IntPtr.Add(block, size), false);
            return new NativeTable { ColumnCount = 2, RowCount = quantities.LongLength, Columns = block };
        }

        private static NativeColumnSpec[] QtyNameSpecs()
        {
            return
            [
                new() { Names = ["qty"], Type = NativeColumnType.Int64 },
                new() { Names = ["name"], Type = NativeColumnType.String },
            ];
        }

        private static NativeTable BuildBoolTable(bool[] values)
        {
            NativeColumn column = new()
            {
                Type = NativeColumnType.Bool,
                Length = values.LongLength,
                Values = Marshal.AllocHGlobal(values.Length),
            };
            for (int index = 0; index < values.Length; index++)
            {
                Marshal.WriteByte(column.Values, index, values[index] ? (byte)1 : (byte)0);
            }
            return SingleColumnTable(column, values.LongLength);
        }

        private static NativeTable BuildFloat64Table(double[] values)
        {
            NativeColumn column = new()
            {
                Type = NativeColumnType.Float64,
                Length = values.LongLength,
                Values = Marshal.AllocHGlobal(values.Length * sizeof(double)),
            };
            Marshal.Copy(values, 0, column.Values, values.Length);
            return SingleColumnTable(column, values.LongLength);
        }

        private static NativeTable BuildDateTable(int[] days)
        {
            NativeColumn column = new()
            {
                Type = NativeColumnType.Date,
                Length = days.LongLength,
                Values = Marshal.AllocHGlobal(days.Length * sizeof(int)),
            };
            Marshal.Copy(days, 0, column.Values, days.Length);
            return SingleColumnTable(column, days.LongLength);
        }

        private static void FreeBuiltTable(ref NativeTable table)
        {
            for (int index = 0; index < table.ColumnCount; index++)
            {
                NativeColumn column = Marshal.PtrToStructure<NativeColumn>(
                    IntPtr.Add(table.Columns, index * Marshal.SizeOf<NativeColumn>()));
                foreach (IntPtr block in new[] { column.Values, column.Validity, column.Data })
                {
                    if (block != IntPtr.Zero)
                    {
                        Marshal.FreeHGlobal(block);
                    }
                }
            }
            Marshal.FreeHGlobal(table.Columns);
            table = default;
        }

        private static NativeWriteOptionsRaw DefaultWriteOptionsRaw()
        {
            return new NativeWriteOptionsRaw { StructSize = Marshal.SizeOf<NativeWriteOptionsRaw>() };
        }

        private static NativeWriteOptions DefaultWriteOptions()
        {
            Assert.True(NativeWriteOptions.TryDecode(DefaultWriteOptionsRaw(), null, out NativeWriteOptions options, out _));
            return options;
        }

        private static void AssertTemporalColumnsCarryANumberFormat(string path)
        {
            using (FileStream file = File.OpenRead(path))
            using (XlsxReader reader = Excel.FromXlsx(file))
            using (XlsxReader.Enumerator rows = reader.GetEnumerator())
            {
                Assert.True(rows.MoveNext());
                Assert.True(rows.MoveNext());
                Assert.Equal(CellType.Date, rows.Current[0].Type);
                Assert.Equal(CellType.Date, rows.Current[2].Type);
                Assert.Equal(CellType.Number, rows.Current[1].Type);
            }

            using ZipArchive archive = ZipFile.OpenRead(path);
            string styles = ReadZipEntry(archive, "xl/styles.xml");
            Assert.Contains("hh:mm:ss", styles, StringComparison.Ordinal);
            Assert.Contains("yyyy-mm-dd hh:mm:ss", styles, StringComparison.Ordinal);
            Assert.Contains("<cols>", ReadZipEntry(archive, "xl/worksheets/sheet1.xml"), StringComparison.Ordinal);
        }

        private static double ReadFirstDataSerial(string path)
        {
            using FileStream file = File.OpenRead(path);
            using XlsbReader reader = Excel.FromXlsb(file);
            using XlsbReader.Enumerator rows = reader.GetEnumerator();

            Assert.True(rows.MoveNext());
            Assert.True(rows.MoveNext());
            Assert.True(rows.Current[0].TryGetDouble(out double serial));
            return serial;
        }

        private static NativeTable BuildTemporalTable(int day, long clock, long stamp)
        {
            NativeColumn[] columns =
            [
                new() { Type = NativeColumnType.Date, Length = 1, Values = Marshal.AllocHGlobal(sizeof(int)) },
                new() { Type = NativeColumnType.Time, Length = 1, Values = Marshal.AllocHGlobal(sizeof(long)) },
                new() { Type = NativeColumnType.Timestamp, Length = 1, Values = Marshal.AllocHGlobal(sizeof(long)) },
            ];
            Marshal.WriteInt32(columns[0].Values, day);
            Marshal.WriteInt64(columns[1].Values, clock);
            Marshal.WriteInt64(columns[2].Values, stamp);

            int size = Marshal.SizeOf<NativeColumn>();
            IntPtr block = Marshal.AllocHGlobal(size * columns.Length);
            for (int index = 0; index < columns.Length; index++)
            {
                Marshal.StructureToPtr(columns[index], IntPtr.Add(block, index * size), false);
            }
            return new NativeTable { ColumnCount = columns.Length, RowCount = 1, Columns = block };
        }

        [Fact]
        public void ValidateWriteTable_Should_Accept_A_Well_Formed_Named_Table()
        {
            NativeTable table = BuildInt64Table([1L, 2L]);
            try
            {
                NativeColumnSpec[] specs = [new() { Names = ["qty"], Type = NativeColumnType.Int64 }];

                Assert.True(WriteApi.TryValidateWriteTable(specs, table, out bool hasHeader, out string? error));
                Assert.True(hasHeader);
                Assert.Null(error);
            }
            finally
            {
                FreeBuiltTable(ref table);
            }
        }

        [Fact]
        public void ValidateWriteTable_Should_Report_No_Header_When_Every_Spec_Is_Unnamed()
        {
            NativeTable table = BuildInt64Table([1L]);
            try
            {
                NativeColumnSpec[] specs = [new() { Names = [], Type = NativeColumnType.Int64 }];

                Assert.True(WriteApi.TryValidateWriteTable(specs, table, out bool hasHeader, out _));
                Assert.False(hasHeader);
            }
            finally
            {
                FreeBuiltTable(ref table);
            }
        }

        [Fact]
        public void ValidateWriteTable_Should_Reject_A_Mix_Of_Named_And_Unnamed_Specs()
        {
            NativeTable table = BuildInt64Table([1L]);
            table.ColumnCount = 2;
            try
            {
                NativeColumnSpec[] specs =
                [
                    new() { Names = ["qty"], Type = NativeColumnType.Int64 },
                    new() { Names = [], Type = NativeColumnType.Int64 },
                ];

                Assert.False(WriteApi.TryValidateWriteTable(specs, table, out _, out string? error));
                Assert.Contains("name", error, StringComparison.Ordinal);
            }
            finally
            {
                table.ColumnCount = 1;
                FreeBuiltTable(ref table);
            }
        }

        [Fact]
        public void ValidateWriteTable_Should_Reject_A_Spec_Type_That_Disagrees_With_Its_Column()
        {
            NativeTable table = BuildInt64Table([1L]);
            try
            {
                NativeColumnSpec[] specs = [new() { Names = ["qty"], Type = NativeColumnType.Float64 }];

                Assert.False(WriteApi.TryValidateWriteTable(specs, table, out _, out string? error));
                Assert.Contains("type", error, StringComparison.Ordinal);
            }
            finally
            {
                FreeBuiltTable(ref table);
            }
        }

        [Fact]
        public void ValidateWriteTable_Should_Reject_A_Column_Whose_Length_Is_Not_The_Row_Count()
        {
            NativeTable table = BuildInt64Table([1L, 2L]);
            table.RowCount = 3;
            try
            {
                NativeColumnSpec[] specs = [new() { Names = ["qty"], Type = NativeColumnType.Int64 }];

                Assert.False(WriteApi.TryValidateWriteTable(specs, table, out _, out string? error));
                Assert.Contains("length", error, StringComparison.Ordinal);
            }
            finally
            {
                table.RowCount = 2;
                FreeBuiltTable(ref table);
            }
        }

        [Fact]
        public void ValidateWriteTable_Should_Accept_Well_Formed_String_Offsets()
        {
            NativeTable table = BuildStringTable([0, 6, 12], "widgetgadget"u8.ToArray());
            try
            {
                NativeColumnSpec[] specs = [new() { Names = ["name"], Type = NativeColumnType.String }];

                Assert.True(WriteApi.TryValidateWriteTable(specs, table, out _, out string? error));
                Assert.Null(error);
            }
            finally
            {
                FreeBuiltTable(ref table);
            }
        }

        [Theory]
        [InlineData(new[] { 1, 6, 12 })]
        [InlineData(new[] { 0, -1, 12 })]
        [InlineData(new[] { 0, 9, 6 })]
        [InlineData(new[] { 0, 6, 13 })]
        [InlineData(new[] { 0, 6, 11 })]
        public void ValidateWriteTable_Should_Reject_Malformed_String_Offsets(int[] offsets)
        {
            NativeTable table = BuildStringTable(offsets, "widgetgadget"u8.ToArray());
            try
            {
                NativeColumnSpec[] specs = [new() { Names = ["name"], Type = NativeColumnType.String }];

                Assert.False(WriteApi.TryValidateWriteTable(specs, table, out _, out string? error));
                Assert.Contains("offset", error, StringComparison.Ordinal);
            }
            finally
            {
                FreeBuiltTable(ref table);
            }
        }

        [Fact]
        public void ValidateWriteTable_Should_Reject_A_Non_Positive_Column_Count()
        {
            NativeTable table = BuildInt64Table([1L]);
            table.ColumnCount = 0;
            try
            {
                Assert.False(WriteApi.TryValidateWriteTable([], table, out _, out string? error));
                Assert.Contains("column", error, StringComparison.Ordinal);
            }
            finally
            {
                table.ColumnCount = 1;
                FreeBuiltTable(ref table);
            }
        }

        [Fact]
        public void ValidateWriteTable_Should_Reject_A_Negative_Row_Count()
        {
            NativeTable table = BuildInt64Table([1L]);
            table.RowCount = -1;
            try
            {
                NativeColumnSpec[] specs = [new() { Names = ["qty"], Type = NativeColumnType.Int64 }];

                Assert.False(WriteApi.TryValidateWriteTable(specs, table, out _, out string? error));
                Assert.Contains("row_count", error, StringComparison.Ordinal);
            }
            finally
            {
                table.RowCount = 1;
                FreeBuiltTable(ref table);
            }
        }

        [Fact]
        public void ValidateWriteTable_Should_Reject_A_Column_Type_Outside_The_Enum()
        {
            NativeTable table = BuildInt64Table([1L]);
            try
            {
                NativeColumn column = Marshal.PtrToStructure<NativeColumn>(table.Columns);
                column.Type = 99;
                Marshal.StructureToPtr(column, table.Columns, false);
                NativeColumnSpec[] specs = [new() { Names = ["qty"], Type = 99 }];

                Assert.False(WriteApi.TryValidateWriteTable(specs, table, out _, out string? error));
                Assert.Contains("unknown type", error, StringComparison.Ordinal);
            }
            finally
            {
                FreeBuiltTable(ref table);
            }
        }

        [Fact]
        public void ValidateWriteTable_Should_Reject_A_Null_Values_Pointer_With_Rows()
        {
            NativeTable table = BuildInt64Table([1L]);
            try
            {
                IntPtr original = Marshal.ReadIntPtr(table.Columns, sizeof(long) * 2);
                NativeColumn column = Marshal.PtrToStructure<NativeColumn>(table.Columns);
                column.Values = IntPtr.Zero;
                Marshal.StructureToPtr(column, table.Columns, false);
                NativeColumnSpec[] specs = [new() { Names = ["qty"], Type = NativeColumnType.Int64 }];

                Assert.False(WriteApi.TryValidateWriteTable(specs, table, out _, out string? error));
                Assert.Contains("values", error, StringComparison.Ordinal);

                column.Values = original;
                Marshal.StructureToPtr(column, table.Columns, false);
            }
            finally
            {
                FreeBuiltTable(ref table);
            }
        }

        [Fact]
        public void WriteOptions_Should_Decode_An_All_Defaults_Struct()
        {
            Assert.True(NativeWriteOptions.TryDecode(DefaultWriteOptionsRaw(), null, out NativeWriteOptions options, out string? error));

            Assert.Null(error);
            Assert.Null(options.SheetName);
            Assert.Null(options.CsvDelimiter);
            Assert.Null(options.CsvQuote);
            Assert.Null(options.Date1904);
            Assert.Null(options.UseSharedStrings);
        }

        [Fact]
        public void WriteOptions_Should_Reject_An_Unrecognized_Struct_Size()
        {
            NativeWriteOptionsRaw raw = DefaultWriteOptionsRaw();
            raw.StructSize = 1;

            Assert.False(NativeWriteOptions.TryDecode(raw, null, out _, out string? error));
            Assert.Contains("struct_size", error, StringComparison.Ordinal);
        }

        [Fact]
        public void WriteOptions_Should_Reject_A_Bad_Struct_Size_Before_Any_Other_Field()
        {
            NativeWriteOptionsRaw raw = DefaultWriteOptionsRaw();
            raw.StructSize = 1;
            raw.SheetNameLen = int.MaxValue;
            raw.CsvDelimiter = 300;

            Assert.False(NativeWriteOptions.TryValidateStructSize(raw, out string? error));
            Assert.Contains("struct_size", error, StringComparison.Ordinal);

            Assert.False(NativeWriteOptions.TryDecode(raw, "has/slash", out _, out string? decodeError));
            Assert.Contains("struct_size", decodeError, StringComparison.Ordinal);
        }

        [Fact]
        public void WriteOptions_Should_Reject_A_Csv_Delimiter_Outside_A_Byte()
        {
            NativeWriteOptionsRaw raw = DefaultWriteOptionsRaw();
            raw.CsvDelimiter = 300;

            Assert.False(NativeWriteOptions.TryDecode(raw, null, out _, out string? error));
            Assert.Contains("csv_delimiter", error, StringComparison.Ordinal);
        }

        [Fact]
        public void WriteOptions_Should_Reject_A_Sheet_Name_Excel_Cannot_Store()
        {
            NativeWriteOptionsRaw raw = DefaultWriteOptionsRaw();

            Assert.False(NativeWriteOptions.TryDecode(raw, "has/slash", out _, out string? error));
            Assert.Contains("sheet_name", error, StringComparison.Ordinal);
            Assert.False(NativeWriteOptions.TryDecode(raw, "", out _, out error));
            Assert.Contains("sheet_name", error, StringComparison.Ordinal);
            Assert.False(NativeWriteOptions.TryDecode(raw, new string('x', 32), out _, out error));
            Assert.Contains("sheet_name", error, StringComparison.Ordinal);
        }

        [Fact]
        public void WriteOptions_Should_Carry_Overrides_Into_CsvWriterOptions()
        {
            NativeWriteOptionsRaw raw = DefaultWriteOptionsRaw();
            raw.CsvDelimiter = ';';
            raw.CsvQuote = '\'';
            raw.Date1904 = NativeOptionState.True;
            raw.UseSharedStrings = NativeOptionState.True;

            Assert.True(NativeWriteOptions.TryDecode(raw, "Data", out NativeWriteOptions options, out _));

            Assert.Equal("Data", options.SheetName);
            Assert.True(options.Date1904);
            Assert.True(options.UseSharedStrings);
            CsvWriterOptions csv = options.ToCsvWriterOptions();
            Assert.Equal((byte)';', csv.Delimiter);
            Assert.Equal((byte)'\'', csv.Quote);
        }

        [Theory]
        [InlineData(NativeFormat.Xlsx, "xlsx")]
        [InlineData(NativeFormat.Xlsb, "xlsb")]
        [InlineData(NativeFormat.Xls, "xls")]
        [InlineData(NativeFormat.Csv, "csv")]
        public void WriteTyped_Should_Round_Trip_Through_ParseTyped(int format, string extension)
        {
            string path = Path.Combine(Path.GetTempPath(), $"excelreader-native-{Guid.NewGuid():N}.{extension}");
            NativeTable table = BuildInt64Table([3L, 7L]);
            try
            {
                NativeColumnSpec[] specs = [new() { Names = ["qty"], Type = NativeColumnType.Int64 }];

                Assert.Equal(NativeStatus.Ok, WriteApi.WriteTyped(
                    Encoding.UTF8.GetBytes(path), format, specs, table, DefaultWriteOptions()));

                Assert.Equal(NativeStatus.Ok, OpenPath(path, format, out NativeHandle? handle));
                try
                {
                    Assert.Equal(NativeStatus.Ok, TypedApi.ParseTyped(handle, specs, headerRow: 1, out NativeTable read));
                    try
                    {
                        Assert.Equal(2, read.RowCount);
                        long[] values = new long[2];
                        Marshal.Copy(ColumnAt(read, 0).Values, values, 0, 2);
                        Assert.Equal([3L, 7L], values);
                    }
                    finally
                    {
                        TypedApi.FreeTable(ref read);
                    }
                }
                finally
                {
                    ReadApi.Close(handle);
                }
            }
            finally
            {
                FreeBuiltTable(ref table);
                File.Delete(path);
            }
        }

        [Theory]
        [InlineData(NativeFormat.Xlsx)]
        [InlineData(NativeFormat.Xlsb)]
        [InlineData(NativeFormat.Xls)]
        [InlineData(NativeFormat.Csv)]
        public void WriteTypedToMemory_Should_Round_Trip_Through_OpenMemory(int format)
        {
            NativeTable table = BuildInt64Table([3L, 7L]);
            try
            {
                NativeColumnSpec[] specs = [new() { Names = ["qty"], Type = NativeColumnType.Int64 }];

                Assert.Equal(NativeStatus.Ok, WriteApi.WriteTypedToMemory(
                    format, specs, table, DefaultWriteOptions(), out byte[]? bytes));
                Assert.NotNull(bytes);
                Assert.NotEmpty(bytes);

                Assert.Equal(NativeStatus.Ok, ReadApi.OpenMemory(bytes, format, out NativeHandle? handle));
                try
                {
                    Assert.Equal(NativeStatus.Ok, TypedApi.ParseTyped(handle, specs, headerRow: 1, out NativeTable read));
                    try
                    {
                        Assert.Equal(2, read.RowCount);
                        long[] values = new long[2];
                        Marshal.Copy(ColumnAt(read, 0).Values, values, 0, 2);
                        Assert.Equal([3L, 7L], values);
                    }
                    finally
                    {
                        TypedApi.FreeTable(ref read);
                    }
                }
                finally
                {
                    ReadApi.Close(handle);
                }
            }
            finally
            {
                FreeBuiltTable(ref table);
            }
        }

        [Fact]
        public void WriteTypedToMemory_Should_Reject_Auto_Format_And_Return_No_Bytes()
        {
            NativeTable table = BuildInt64Table([3L]);
            try
            {
                NativeColumnSpec[] specs = [new() { Names = ["qty"], Type = NativeColumnType.Int64 }];

                Assert.Equal(NativeStatus.InvalidArgument, WriteApi.WriteTypedToMemory(
                    NativeFormat.Auto, specs, table, DefaultWriteOptions(), out byte[]? bytes));
                Assert.Null(bytes);
            }
            finally
            {
                FreeBuiltTable(ref table);
            }
        }

        [Fact]
        public void WriteTypedToMemory_Should_Reject_A_Rejected_Table_The_Same_Way_As_WriteTyped()
        {
            NativeTable table = BuildInt64Table([3L]);
            try
            {
                NativeColumnSpec[] specs =
                [
                    new() { Names = ["a"], Type = NativeColumnType.Int64 },
                    new() { Names = ["b"], Type = NativeColumnType.Int64 },
                ];

                Assert.Equal(NativeStatus.InvalidArgument, WriteApi.WriteTypedToMemory(
                    NativeFormat.Xlsx, specs, table, DefaultWriteOptions(), out byte[]? bytes));
                Assert.Null(bytes);
            }
            finally
            {
                FreeBuiltTable(ref table);
            }
        }

        [Theory]
        [InlineData(NativeFormat.Xlsx, "xlsx")]
        [InlineData(NativeFormat.Xlsb, "xlsb")]
        [InlineData(NativeFormat.Xls, "xls")]
        [InlineData(NativeFormat.Csv, "csv")]
        public void WriteTyped_Should_Round_Trip_Bools_Through_ParseTyped(int format, string extension)
        {
            string path = Path.Combine(Path.GetTempPath(), $"excelreader-native-{Guid.NewGuid():N}.{extension}");
            NativeTable table = BuildBoolTable([true, false, true]);
            try
            {
                NativeColumnSpec[] specs = [new() { Names = ["flag"], Type = NativeColumnType.Bool }];

                Assert.Equal(NativeStatus.Ok, WriteApi.WriteTyped(
                    Encoding.UTF8.GetBytes(path), format, specs, table, DefaultWriteOptions()));

                Assert.Equal(NativeStatus.Ok, OpenPath(path, format, out NativeHandle? handle));
                try
                {
                    Assert.Equal(NativeStatus.Ok, TypedApi.ParseTyped(handle, specs, headerRow: 1, out NativeTable read));
                    try
                    {
                        Assert.Equal(3, read.RowCount);
                        NativeColumn column = ColumnAt(read, 0);
                        byte[] flags = [Marshal.ReadByte(column.Values, 0), Marshal.ReadByte(column.Values, 1), Marshal.ReadByte(column.Values, 2)];
                        Assert.Equal<byte>([1, 0, 1], flags);
                    }
                    finally
                    {
                        TypedApi.FreeTable(ref read);
                    }
                }
                finally
                {
                    ReadApi.Close(handle);
                }
            }
            finally
            {
                FreeBuiltTable(ref table);
                File.Delete(path);
            }
        }

        [Theory]
        [InlineData(NativeFormat.Xlsx, "xlsx")]
        [InlineData(NativeFormat.Xlsb, "xlsb")]
        [InlineData(NativeFormat.Xls, "xls")]
        [InlineData(NativeFormat.Csv, "csv")]
        public void WriteTyped_Should_Round_Trip_Doubles_Through_ParseTyped(int format, string extension)
        {
            string path = Path.Combine(Path.GetTempPath(), $"excelreader-native-{Guid.NewGuid():N}.{extension}");
            NativeTable table = BuildFloat64Table([3.5, -0.25]);
            try
            {
                NativeColumnSpec[] specs = [new() { Names = ["rate"], Type = NativeColumnType.Float64 }];

                Assert.Equal(NativeStatus.Ok, WriteApi.WriteTyped(
                    Encoding.UTF8.GetBytes(path), format, specs, table, DefaultWriteOptions()));

                Assert.Equal(NativeStatus.Ok, OpenPath(path, format, out NativeHandle? handle));
                try
                {
                    Assert.Equal(NativeStatus.Ok, TypedApi.ParseTyped(handle, specs, headerRow: 1, out NativeTable read));
                    try
                    {
                        Assert.Equal(2, read.RowCount);
                        double[] values = new double[2];
                        Marshal.Copy(ColumnAt(read, 0).Values, values, 0, 2);
                        Assert.Equal([3.5, -0.25], values);
                    }
                    finally
                    {
                        TypedApi.FreeTable(ref read);
                    }
                }
                finally
                {
                    ReadApi.Close(handle);
                }
            }
            finally
            {
                FreeBuiltTable(ref table);
                File.Delete(path);
            }
        }

        [Fact]
        public void WriteTyped_Should_Round_Trip_Strings_And_Nulls()
        {
            string path = Path.Combine(Path.GetTempPath(), $"excelreader-native-{Guid.NewGuid():N}.xlsx");
            NativeTable table = BuildStringTable([0, 0, 6, 12, 12], "widgetgadget"u8.ToArray());
            try
            {
                NativeColumn column = Marshal.PtrToStructure<NativeColumn>(table.Columns);
                column.Validity = Marshal.AllocHGlobal(1);
                Marshal.WriteByte(column.Validity, 0b0110);
                Marshal.StructureToPtr(column, table.Columns, false);

                NativeColumnSpec[] specs = [new() { Names = ["name"], Type = NativeColumnType.String }];
                Assert.Equal(NativeStatus.Ok, WriteApi.WriteTyped(
                    Encoding.UTF8.GetBytes(path), NativeFormat.Xlsx, specs, table, DefaultWriteOptions()));

                Assert.Equal(NativeStatus.Ok, OpenPath(path, NativeFormat.Xlsx, out NativeHandle? handle));
                try
                {
                    NativeColumnSpec[] readSpecs = [new() { Names = ["name"], Type = NativeColumnType.String, Nullable = true }];
                    Assert.Equal(NativeStatus.Ok, TypedApi.ParseTyped(handle, readSpecs, headerRow: 1, out NativeTable read));
                    try
                    {
                        Assert.Equal(4, read.RowCount);
                        Assert.Equal(["", "widget", "gadget", ""], DecodeStringColumn(ColumnAt(read, 0)));
                        Assert.Equal([true, true, true, true], DecodeValidity(ColumnAt(read, 0)));
                    }
                    finally
                    {
                        TypedApi.FreeTable(ref read);
                    }
                }
                finally
                {
                    ReadApi.Close(handle);
                }
            }
            finally
            {
                FreeBuiltTable(ref table);
                File.Delete(path);
            }
        }

        [Fact]
        public void WriteTyped_Should_Round_Trip_A_Validity_Bitmap_With_Nulls_At_Both_Ends()
        {
            string path = Path.Combine(Path.GetTempPath(), $"excelreader-native-{Guid.NewGuid():N}.xlsx");
            NativeTable table = BuildInt64Table([0L, 7L, 0L], [0b010]);
            try
            {
                NativeColumnSpec[] specs = [new() { Names = ["qty"], Type = NativeColumnType.Int64, Nullable = true }];
                Assert.Equal(NativeStatus.Ok, WriteApi.WriteTyped(
                    Encoding.UTF8.GetBytes(path), NativeFormat.Xlsx, specs, table, DefaultWriteOptions()));

                Assert.Equal(NativeStatus.Ok, OpenPath(path, NativeFormat.Xlsx, out NativeHandle? handle));
                try
                {
                    Assert.Equal(NativeStatus.Ok, TypedApi.ParseTyped(handle, specs, headerRow: 1, out NativeTable read));
                    try
                    {
                        Assert.Equal(3, read.RowCount);
                        Assert.Equal([false, true, false], DecodeValidity(ColumnAt(read, 0)));
                        long[] values = new long[3];
                        Marshal.Copy(ColumnAt(read, 0).Values, values, 0, 3);
                        Assert.Equal(7L, values[1]);
                    }
                    finally
                    {
                        TypedApi.FreeTable(ref read);
                    }
                }
                finally
                {
                    ReadApi.Close(handle);
                }
            }
            finally
            {
                FreeBuiltTable(ref table);
                File.Delete(path);
            }
        }

        [Fact]
        public void WriteTyped_Should_Round_Trip_A_String_Column_That_Is_Entirely_Empty()
        {
            string path = Path.Combine(Path.GetTempPath(), $"excelreader-native-{Guid.NewGuid():N}.xlsx");
            NativeTable table = BuildStringTable([0, 0, 0], []);
            try
            {
                NativeColumnSpec[] specs = [new() { Names = ["name"], Type = NativeColumnType.String }];

                Assert.Equal(NativeStatus.Ok, WriteApi.WriteTyped(
                    Encoding.UTF8.GetBytes(path), NativeFormat.Xlsx, specs, table, DefaultWriteOptions()));

                Assert.Equal(NativeStatus.Ok, OpenPath(path, NativeFormat.Xlsx, out NativeHandle? handle));
                try
                {
                    Assert.Equal(NativeStatus.Ok, TypedApi.ParseTyped(handle, specs, headerRow: 1, out NativeTable read));
                    try
                    {
                        Assert.Equal(2, read.RowCount);
                        Assert.Equal(["", ""], DecodeStringColumn(ColumnAt(read, 0)));
                    }
                    finally
                    {
                        TypedApi.FreeTable(ref read);
                    }
                }
                finally
                {
                    ReadApi.Close(handle);
                }
            }
            finally
            {
                FreeBuiltTable(ref table);
                File.Delete(path);
            }
        }

        [Fact]
        public void WriteTyped_Should_Round_Trip_Every_Temporal_Type()
        {
            string path = Path.Combine(Path.GetTempPath(), $"excelreader-native-{Guid.NewGuid():N}.xlsx");
            int epoch = new DateOnly(1970, 1, 1).DayNumber;
            int day = new DateOnly(2024, 1, 15).DayNumber - epoch;
            long clock = 3_600_000_000L;
            long stamp = 1_705_280_400_000_000L;

            NativeTable table = BuildTemporalTable(day, clock, stamp);
            try
            {
                NativeColumnSpec[] specs =
                [
                    new() { Names = ["day"], Type = NativeColumnType.Date },
                    new() { Names = ["clock"], Type = NativeColumnType.Time },
                    new() { Names = ["stamp"], Type = NativeColumnType.Timestamp },
                ];
                Assert.Equal(NativeStatus.Ok, WriteApi.WriteTyped(
                    Encoding.UTF8.GetBytes(path), NativeFormat.Xlsx, specs, table, DefaultWriteOptions()));

                Assert.Equal(NativeStatus.Ok, OpenPath(path, NativeFormat.Xlsx, out NativeHandle? handle));
                try
                {
                    Assert.Equal(NativeStatus.Ok, TypedApi.ParseTyped(handle, specs, headerRow: 1, out NativeTable read));
                    try
                    {
                        int[] days = new int[1];
                        Marshal.Copy(ColumnAt(read, 0).Values, days, 0, 1);
                        long[] clocks = new long[1];
                        Marshal.Copy(ColumnAt(read, 1).Values, clocks, 0, 1);
                        long[] stamps = new long[1];
                        Marshal.Copy(ColumnAt(read, 2).Values, stamps, 0, 1);

                        Assert.Equal(day, days[0]);
                        Assert.Equal(clock, clocks[0]);
                        Assert.Equal(stamp, stamps[0]);
                    }
                    finally
                    {
                        TypedApi.FreeTable(ref read);
                    }
                }
                finally
                {
                    ReadApi.Close(handle);
                }

                AssertTemporalColumnsCarryANumberFormat(path);
            }
            finally
            {
                FreeBuiltTable(ref table);
                File.Delete(path);
            }
        }

        [Fact]
        public void WriteTyped_Should_Carry_Date1904_Into_The_Written_Xlsb()
        {
            string path = Path.Combine(Path.GetTempPath(), $"excelreader-native-{Guid.NewGuid():N}.xlsb");
            int day = new DateOnly(2024, 1, 15).DayNumber - new DateOnly(1970, 1, 1).DayNumber;
            NativeTable table = BuildDateTable([day]);
            try
            {
                NativeWriteOptionsRaw raw = DefaultWriteOptionsRaw();
                raw.Date1904 = NativeOptionState.True;
                Assert.True(NativeWriteOptions.TryDecode(raw, null, out NativeWriteOptions options, out _));
                NativeColumnSpec[] specs = [new() { Names = ["day"], Type = NativeColumnType.Date }];

                Assert.Equal(NativeStatus.Ok, WriteApi.WriteTyped(
                    Encoding.UTF8.GetBytes(path), NativeFormat.Xlsb, specs, table, options));

                Assert.Equal(NativeStatus.Ok, OpenPath(path, NativeFormat.Xlsb, out NativeHandle? handle));
                try
                {
                    Assert.Equal(NativeStatus.Ok, ReadApi.IsDate1904(handle, out int flag));
                    Assert.Equal(1, flag);

                    Assert.Equal(NativeStatus.Ok, TypedApi.ParseTyped(handle, specs, headerRow: 1, out NativeTable read));
                    try
                    {
                        int[] days = new int[1];
                        Marshal.Copy(ColumnAt(read, 0).Values, days, 0, 1);
                        Assert.Equal(day, days[0]);
                    }
                    finally
                    {
                        TypedApi.FreeTable(ref read);
                    }
                }
                finally
                {
                    ReadApi.Close(handle);
                }
            }
            finally
            {
                FreeBuiltTable(ref table);
                File.Delete(path);
            }
        }

        [Fact]
        public void WriteTyped_Should_Store_A_Different_Serial_Under_The_1904_Epoch()
        {
            string epoch1900 = Path.Combine(Path.GetTempPath(), $"excelreader-native-{Guid.NewGuid():N}.xlsb");
            string epoch1904 = Path.Combine(Path.GetTempPath(), $"excelreader-native-{Guid.NewGuid():N}.xlsb");
            int day = new DateOnly(2024, 1, 15).DayNumber - new DateOnly(1970, 1, 1).DayNumber;
            NativeTable table = BuildDateTable([day]);
            try
            {
                NativeWriteOptionsRaw raw = DefaultWriteOptionsRaw();
                raw.Date1904 = NativeOptionState.True;
                Assert.True(NativeWriteOptions.TryDecode(raw, null, out NativeWriteOptions options1904, out _));
                NativeColumnSpec[] specs = [new() { Names = ["day"], Type = NativeColumnType.Date }];

                Assert.Equal(NativeStatus.Ok, WriteApi.WriteTyped(
                    Encoding.UTF8.GetBytes(epoch1900), NativeFormat.Xlsb, specs, table, DefaultWriteOptions()));
                Assert.Equal(NativeStatus.Ok, WriteApi.WriteTyped(
                    Encoding.UTF8.GetBytes(epoch1904), NativeFormat.Xlsb, specs, table, options1904));

                Assert.Equal(1462.0, ReadFirstDataSerial(epoch1900) - ReadFirstDataSerial(epoch1904));
            }
            finally
            {
                FreeBuiltTable(ref table);
                File.Delete(epoch1900);
                File.Delete(epoch1904);
            }
        }

        [Fact]
        public void WriteTyped_Should_Write_No_Header_Row_When_Specs_Are_Unnamed()
        {
            string path = Path.Combine(Path.GetTempPath(), $"excelreader-native-{Guid.NewGuid():N}.csv");
            NativeTable table = BuildInt64Table([3L, 7L]);
            try
            {
                NativeColumnSpec[] specs = [new() { Names = [], Type = NativeColumnType.Int64 }];
                Assert.Equal(NativeStatus.Ok, WriteApi.WriteTyped(
                    Encoding.UTF8.GetBytes(path), NativeFormat.Csv, specs, table, DefaultWriteOptions()));

                Assert.Equal("3\n7\n", File.ReadAllText(path).ReplaceLineEndings("\n"));
            }
            finally
            {
                FreeBuiltTable(ref table);
                File.Delete(path);
            }
        }

        [Fact]
        public void WriteTyped_Should_Apply_The_Csv_Delimiter_Override()
        {
            string path = Path.Combine(Path.GetTempPath(), $"excelreader-native-{Guid.NewGuid():N}.csv");
            NativeTable table = BuildQtyNameTable([3L], ["widget"]);
            try
            {
                NativeWriteOptionsRaw raw = DefaultWriteOptionsRaw();
                raw.CsvDelimiter = ';';
                Assert.True(NativeWriteOptions.TryDecode(raw, null, out NativeWriteOptions options, out _));
                NativeColumnSpec[] specs = QtyNameSpecs();

                Assert.Equal(NativeStatus.Ok, WriteApi.WriteTyped(
                    Encoding.UTF8.GetBytes(path), NativeFormat.Csv, specs, table, options));

                Assert.Equal("qty;name\n3;widget\n", File.ReadAllText(path).ReplaceLineEndings("\n"));
            }
            finally
            {
                FreeBuiltTable(ref table);
                File.Delete(path);
            }
        }

        [Fact]
        public void WriteTyped_Should_Apply_The_Csv_Quote_Override()
        {
            string path = Path.Combine(Path.GetTempPath(), $"excelreader-native-{Guid.NewGuid():N}.csv");
            NativeTable table = BuildQtyNameTable([3L], ["wid,get"]);
            try
            {
                NativeWriteOptionsRaw raw = DefaultWriteOptionsRaw();
                raw.CsvQuote = '\'';
                Assert.True(NativeWriteOptions.TryDecode(raw, null, out NativeWriteOptions options, out _));
                NativeColumnSpec[] specs = QtyNameSpecs();

                Assert.Equal(NativeStatus.Ok, WriteApi.WriteTyped(
                    Encoding.UTF8.GetBytes(path), NativeFormat.Csv, specs, table, options));

                Assert.Equal("qty,name\n3,'wid,get'\n", File.ReadAllText(path).ReplaceLineEndings("\n"));
            }
            finally
            {
                FreeBuiltTable(ref table);
                File.Delete(path);
            }
        }

        [Fact]
        public void WriteTyped_Should_Reject_A_Spec_With_More_Than_One_Name()
        {
            NativeTable table = BuildInt64Table([1L]);
            try
            {
                NativeColumnSpec[] specs = [new() { Names = ["qty", "quantity"], Type = NativeColumnType.Int64 }];
                string path = Path.Combine(Path.GetTempPath(), $"excelreader-native-{Guid.NewGuid():N}.csv");
                try
                {
                    int status = WriteApi.WriteTyped(Encoding.UTF8.GetBytes(path), NativeFormat.Csv, specs, table, new NativeWriteOptions());
                    Assert.Equal(NativeStatus.InvalidArgument, status);
                    Assert.False(File.Exists(path));
                }
                finally
                {
                    if (File.Exists(path))
                    {
                        File.Delete(path);
                    }
                }
            }
            finally
            {
                FreeBuiltTable(ref table);
            }
        }

        [Fact]
        public void WriteTyped_Should_Reject_Auto_Format_And_Create_No_File()
        {
            string path = Path.Combine(Path.GetTempPath(), $"excelreader-native-{Guid.NewGuid():N}.xlsx");
            NativeTable table = BuildInt64Table([3L]);
            try
            {
                NativeColumnSpec[] specs = [new() { Names = ["qty"], Type = NativeColumnType.Int64 }];

                Assert.Equal(NativeStatus.InvalidArgument, WriteApi.WriteTyped(
                    Encoding.UTF8.GetBytes(path), NativeFormat.Auto, specs, table, DefaultWriteOptions()));
                Assert.False(File.Exists(path));
            }
            finally
            {
                FreeBuiltTable(ref table);
                File.Delete(path);
            }
        }

        [Fact]
        public void WriteTyped_Should_Reject_A_Format_Outside_Every_Constant()
        {
            string path = Path.Combine(Path.GetTempPath(), $"excelreader-native-{Guid.NewGuid():N}.xlsx");
            NativeTable table = BuildInt64Table([3L]);
            try
            {
                NativeColumnSpec[] specs = [new() { Names = ["qty"], Type = NativeColumnType.Int64 }];

                Assert.Equal(NativeStatus.InvalidArgument, WriteApi.WriteTyped(
                    Encoding.UTF8.GetBytes(path), format: 99, specs, table, DefaultWriteOptions()));
                Assert.False(File.Exists(path));
            }
            finally
            {
                FreeBuiltTable(ref table);
                File.Delete(path);
            }
        }

        [Fact]
        public void WriteTyped_Should_Create_No_File_When_The_Table_Is_Rejected()
        {
            string path = Path.Combine(Path.GetTempPath(), $"excelreader-native-{Guid.NewGuid():N}.xlsx");
            NativeTable table = BuildStringTable([0, 9, 6], "widgetgadget"u8.ToArray());
            try
            {
                NativeColumnSpec[] specs = [new() { Names = ["name"], Type = NativeColumnType.String }];

                Assert.Equal(NativeStatus.InvalidArgument, WriteApi.WriteTyped(
                    Encoding.UTF8.GetBytes(path), NativeFormat.Xlsx, specs, table, DefaultWriteOptions()));
                Assert.False(File.Exists(path));
            }
            finally
            {
                FreeBuiltTable(ref table);
                File.Delete(path);
            }
        }

        [Fact]
        public void WriteTyped_Should_Use_The_Requested_Sheet_Name()
        {
            string path = Path.Combine(Path.GetTempPath(), $"excelreader-native-{Guid.NewGuid():N}.xlsx");
            NativeTable table = BuildInt64Table([3L]);
            try
            {
                Assert.True(NativeWriteOptions.TryDecode(DefaultWriteOptionsRaw(), "Vendas", out NativeWriteOptions options, out _));
                NativeColumnSpec[] specs = [new() { Names = ["qty"], Type = NativeColumnType.Int64 }];

                Assert.Equal(NativeStatus.Ok, WriteApi.WriteTyped(
                    Encoding.UTF8.GetBytes(path), NativeFormat.Xlsx, specs, table, options));

                Assert.Equal(NativeStatus.Ok, OpenPath(path, NativeFormat.Xlsx, out NativeHandle? handle));
                try
                {
                    byte[] name = new byte[64];
                    Assert.Equal(NativeStatus.Ok, ReadApi.SheetNameAt(handle, 0, name, out int written));
                    Assert.Equal("Vendas", Encoding.UTF8.GetString(name, 0, written));
                }
                finally
                {
                    ReadApi.Close(handle);
                }
            }
            finally
            {
                FreeBuiltTable(ref table);
                File.Delete(path);
            }
        }
    }
}
