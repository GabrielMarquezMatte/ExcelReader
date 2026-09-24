using System.Runtime.InteropServices;
using System.Text;
using ExcelReader.Native;
using ExcelReader.Native.Arrow;
using ExcelReader.Native.Reading;

namespace ExcelReader.Tests.Native
{
    public sealed partial class NativeApiTests
    {
        private static ArrowSchema ArrowChildSchema(ArrowSchema schema, int index)
        {
            return Marshal.PtrToStructure<ArrowSchema>(Marshal.ReadIntPtr(schema.Children, index * IntPtr.Size));
        }

        internal static ArrowArray ArrowChildArray(ArrowArray array, int index)
        {
            return Marshal.PtrToStructure<ArrowArray>(Marshal.ReadIntPtr(array.Children, index * IntPtr.Size));
        }

        internal static IntPtr ArrowBuffer(ArrowArray array, int index)
        {
            return Marshal.ReadIntPtr(array.Buffers, index * IntPtr.Size);
        }

        internal static void ExercisedReleaseArrow(ref ArrowArray array, ref ArrowSchema schema)
        {
            IntPtr arrayBlock = Marshal.AllocHGlobal(Marshal.SizeOf<ArrowArray>());
            IntPtr schemaBlock = Marshal.AllocHGlobal(Marshal.SizeOf<ArrowSchema>());
            try
            {
                Marshal.StructureToPtr(array, arrayBlock, false);
                Marshal.StructureToPtr(schema, schemaBlock, false);
                ArrowApi.ReleaseArrowArray(arrayBlock);
                ArrowApi.ReleaseArrowSchema(schemaBlock);
                array = Marshal.PtrToStructure<ArrowArray>(arrayBlock);
                schema = Marshal.PtrToStructure<ArrowSchema>(schemaBlock);
            }
            finally
            {
                Marshal.FreeHGlobal(arrayBlock);
                Marshal.FreeHGlobal(schemaBlock);
            }
        }

        [Fact]
        public void ParseArrow_Should_Return_A_Struct_Array_With_A_Matching_Schema()
        {
            string path = Path.Combine(Path.GetTempPath(), $"excelreader-native-{Guid.NewGuid():N}.csv");
            File.WriteAllText(path, "name,qty\nwidget,3\ngadget,7\n");
            try
            {
                Assert.Equal(NativeStatus.Ok, OpenPath(path, NativeFormat.Csv, out NativeHandle? handle));
                NativeColumnSpec[] specs =
                [
                    new() { Names = ["name"], Type = NativeColumnType.String },
                    new() { Names = ["qty"], Type = NativeColumnType.Int64 },
                ];
                Assert.Equal(NativeStatus.Ok, ArrowApi.ParseArrow(handle, specs, headerRow: 1, out ArrowArray array, out ArrowSchema schema));
                try
                {
                    Assert.Equal("+s", Marshal.PtrToStringUTF8(schema.Format));
                    Assert.Equal(2, schema.NChildren);
                    Assert.Equal("u", Marshal.PtrToStringUTF8(ArrowChildSchema(schema, 0).Format));
                    Assert.Equal("name", Marshal.PtrToStringUTF8(ArrowChildSchema(schema, 0).Name));
                    Assert.Equal("l", Marshal.PtrToStringUTF8(ArrowChildSchema(schema, 1).Format));

                    Assert.Equal(2, array.Length);
                    Assert.Equal(2, array.NChildren);
                    Assert.Equal(1, array.NBuffers);
                    Assert.Equal(IntPtr.Zero, ArrowBuffer(array, 0));

                    ArrowArray nameColumn = ArrowChildArray(array, 0);
                    Assert.Equal(2, nameColumn.Length);
                    Assert.Equal(0, nameColumn.NullCount);
                    Assert.Equal(3, nameColumn.NBuffers);
                    Assert.Equal(IntPtr.Zero, ArrowBuffer(nameColumn, 0));
                    int[] offsets = new int[3];
                    Marshal.Copy(ArrowBuffer(nameColumn, 1), offsets, 0, 3);
                    byte[] data = new byte[offsets[2]];
                    Marshal.Copy(ArrowBuffer(nameColumn, 2), data, 0, data.Length);
                    Assert.Equal("widget", Encoding.UTF8.GetString(data, offsets[0], offsets[1] - offsets[0]));
                    Assert.Equal("gadget", Encoding.UTF8.GetString(data, offsets[1], offsets[2] - offsets[1]));

                    ArrowArray qtyColumn = ArrowChildArray(array, 1);
                    Assert.Equal(2, qtyColumn.NBuffers);
                    long[] qty = new long[2];
                    Marshal.Copy(ArrowBuffer(qtyColumn, 1), qty, 0, 2);
                    Assert.Equal([3L, 7L], qty);
                }
                finally
                {
                    ExercisedReleaseArrow(ref array, ref schema);
                    ReadApi.Close(handle);
                }
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void ParseArrow_Should_Bit_Pack_Bool_Columns()
        {
            string path = Path.Combine(Path.GetTempPath(), $"excelreader-native-{Guid.NewGuid():N}.csv");
            File.WriteAllText(path, "flag\ntrue\nfalse\ntrue\nfalse\ntrue\nfalse\ntrue\nfalse\ntrue\ntrue\n");
            try
            {
                Assert.Equal(NativeStatus.Ok, OpenPath(path, NativeFormat.Csv, out NativeHandle? handle));
                NativeColumnSpec[] specs = [new() { Names = ["flag"], Type = NativeColumnType.Bool }];
                Assert.Equal(NativeStatus.Ok, ArrowApi.ParseArrow(handle, specs, headerRow: 1, out ArrowArray array, out ArrowSchema schema));
                try
                {
                    ArrowArray column = ArrowChildArray(array, 0);
                    Assert.Equal(10, column.Length);
                    Assert.Equal("b", Marshal.PtrToStringUTF8(ArrowChildSchema(schema, 0).Format));

                    byte[] bitmap = new byte[2];
                    Marshal.Copy(ArrowBuffer(column, 1), bitmap, 0, 2);
                    bool[] expected = [true, false, true, false, true, false, true, false, true, true];
                    for (int i = 0; i < expected.Length; i++)
                    {
                        bool bit = (bitmap[i >> 3] & (1 << (i & 7))) != 0;
                        Assert.Equal(expected[i], bit);
                    }
                }
                finally
                {
                    ExercisedReleaseArrow(ref array, ref schema);
                    ReadApi.Close(handle);
                }
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void ParseArrow_Should_Report_Null_Count_From_The_Validity_Bitmap()
        {
            string path = Path.Combine(Path.GetTempPath(), $"excelreader-native-{Guid.NewGuid():N}.csv");
            File.WriteAllText(path, "qty\n5\n\nnotanumber\n");
            try
            {
                Assert.Equal(NativeStatus.Ok, OpenPath(path, NativeFormat.Csv, out NativeHandle? handle));
                NativeColumnSpec[] specs = [new() { Names = ["qty"], Type = NativeColumnType.Int64, Nullable = true }];
                Assert.Equal(NativeStatus.Ok, ArrowApi.ParseArrow(handle, specs, headerRow: 1, out ArrowArray array, out ArrowSchema schema));
                try
                {
                    ArrowArray column = ArrowChildArray(array, 0);
                    Assert.NotEqual(IntPtr.Zero, ArrowBuffer(column, 0));
                    Assert.Equal(2, column.NullCount);
                    Assert.True((ArrowChildSchema(schema, 0).Flags & ArrowFlags.Nullable) != 0);
                }
                finally
                {
                    ExercisedReleaseArrow(ref array, ref schema);
                    ReadApi.Close(handle);
                }
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void ParseArrow_Should_Propagate_A_Non_Nullable_Conversion_Failure()
        {
            string path = Path.Combine(Path.GetTempPath(), $"excelreader-native-{Guid.NewGuid():N}.csv");
            File.WriteAllText(path, "qty\nnotanumber\n");
            try
            {
                Assert.Equal(NativeStatus.Ok, OpenPath(path, NativeFormat.Csv, out NativeHandle? handle));
                NativeColumnSpec[] specs = [new() { Names = ["qty"], Type = NativeColumnType.Int64 }];

                int status = ArrowApi.ParseArrow(handle, specs, headerRow: 1, out ArrowArray array, out ArrowSchema schema);

                Assert.Equal(NativeStatus.Error, status);
                Assert.Equal(IntPtr.Zero, array.Release);
                Assert.Equal(IntPtr.Zero, schema.Release);
                ReadApi.Close(handle);
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void ParseArrow_Should_Reject_A_Null_Handle()
        {
            NativeColumnSpec[] specs = [new() { Index = 0, Type = NativeColumnType.String }];
            Assert.Equal(NativeStatus.InvalidHandle, ArrowApi.ParseArrow(null, specs, headerRow: 1, out _, out _));
        }

        [Fact]
        public void ReleaseArrowArray_And_ReleaseArrowSchema_Are_Idempotent()
        {
            string path = Path.Combine(Path.GetTempPath(), $"excelreader-native-{Guid.NewGuid():N}.csv");
            File.WriteAllText(path, "name\nwidget\n");
            try
            {
                Assert.Equal(NativeStatus.Ok, OpenPath(path, NativeFormat.Csv, out NativeHandle? handle));
                NativeColumnSpec[] specs = [new() { Names = ["name"], Type = NativeColumnType.String }];
                Assert.Equal(NativeStatus.Ok, ArrowApi.ParseArrow(handle, specs, headerRow: 1, out ArrowArray array, out ArrowSchema schema));

                ExercisedReleaseArrow(ref array, ref schema);
                ExercisedReleaseArrow(ref array, ref schema);

                Assert.Equal(IntPtr.Zero, array.Release);
                Assert.Equal(IntPtr.Zero, schema.Release);
                ReadApi.Close(handle);
            }
            finally
            {
                File.Delete(path);
            }
        }
    }
}
