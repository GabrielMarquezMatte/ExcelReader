using System.Runtime.InteropServices;
using System.Text;
using ExcelReader.Core.Reader.Columnar;
using ExcelReader.Core.Writer.Xlsb;
using ExcelReader.Native;
using ExcelReader.Native.Arrow;
using ExcelReader.Native.Reading;
using ExcelReader.Native.Typed;

namespace ExcelReader.Tests.Native
{
    public sealed partial class NativeApiTests
    {
        private static readonly string[] Cities = ["Lisboa", "Porto", "Braga", "Faro", "Évora"];

        private static string? CityAt(int row)
        {
            return row % 7 == 3 ? null : Cities[row % Cities.Length];
        }

        private static byte[] CityXlsb(int rows, bool sharedStrings)
        {
            using MemoryStream ms = new();
            using (XlsbWorkbookWriter workbook = XlsbWorkbookWriter.Create(ms, leaveOpen: true, new XlsbWriterOptions { UseSharedStrings = sharedStrings }))
            {
                using XlsbSheetWriter sheet = workbook.AddSheet("S");
                using (XlsbRowWriter header = sheet.StartRow())
                {
                    header.Write("city");
                }
                for (int i = 0; i < rows; i++)
                {
                    using XlsbRowWriter row = sheet.StartRow();
                    WriteCity(row, i);
                }
                sheet.End();
                workbook.End();
            }
            return ms.ToArray();
        }

        private static void WriteCity(XlsbRowWriter row, int index)
        {
            if (index % 11 == 5)
            {
                row.Write(42);
                return;
            }
            row.Write(CityAt(index));
        }

        private static string? ExpectedCity(int row)
        {
            return row % 11 == 5 ? "42" : CityAt(row) ?? "";
        }

        private static string?[] DecodeDictionary(in BuiltColumn column)
        {
            int entries = (int)column.DictionaryLength;
            int[] offsets = new int[entries + 1];
            Marshal.Copy(column.DictionaryOffsets, offsets, 0, offsets.Length);
            byte[] data = new byte[column.DataLen];
            Marshal.Copy(column.Data, data, 0, data.Length);
            int[] indices = new int[column.Length];
            Marshal.Copy(column.Values, indices, 0, indices.Length);
            string?[] values = new string?[indices.Length];
            for (int i = 0; i < indices.Length; i++)
            {
                values[i] = Encoding.UTF8.GetString(data, offsets[indices[i]], offsets[indices[i] + 1] - offsets[indices[i]]);
            }
            return values;
        }

        [Theory]
        [InlineData(true, 1, 0)]
        [InlineData(false, 1, 0)]
        [InlineData(true, 4, 16_384)]
        [InlineData(false, 4, 16_384)]
        public void ParseColumns_Should_Dictionary_Encode_Shared_Inline_And_Number_Cells(bool sharedStrings, int dop, int chunkBytes)
        {
            const int Rows = 30_000;
            Assert.Equal(NativeStatus.Ok, ReadApi.OpenMemory(CityXlsb(Rows, sharedStrings), NativeFormat.Xlsb, out NativeHandle? handle));
            try
            {
                NativeColumnSpec[] specs = [new() { Names = ["city"], Type = NativeColumnType.StringDictionary }];
                Assert.Equal(NativeStatus.Ok, TypedApi.ParseColumns(handle, 0, specs, 1, dop, allowDictionary: true, chunkBytes, out ColumnarTable? table));
                using ColumnarTable built = table!;
                Assert.Equal(dop > 1, TypedApi.LastParseRanInParallel);

                BuiltColumn column = built.Columns[0];
                Assert.Equal(Rows, column.Length);
                Assert.Equal(Cities.Length + 2, column.DictionaryLength);
                Assert.Equal(Enumerable.Range(0, Rows).Select(ExpectedCity), DecodeDictionary(in column), StringComparer.Ordinal);
            }
            finally
            {
                ReadApi.Close(handle);
            }
        }

        [Fact]
        public void ParseArrow_Should_Emit_A_Dictionary_Column_With_A_Utf8_Dictionary()
        {
            Assert.Equal(NativeStatus.Ok, ReadApi.OpenMemory(CityXlsb(20, sharedStrings: true), NativeFormat.Xlsb, out NativeHandle? handle));
            NativeColumnSpec[] specs = [new() { Names = ["city"], Type = NativeColumnType.StringDictionary, Nullable = true }];
            Assert.Equal(NativeStatus.Ok, ArrowApi.ParseArrow(handle, 0, specs, headerRow: 1, out ArrowArray array, out ArrowSchema schema));
            try
            {
                ArrowSchema child = ArrowChildSchema(schema, 0);
                Assert.Equal("i", Marshal.PtrToStringUTF8(child.Format));
                Assert.Equal("u", Marshal.PtrToStringUTF8(Marshal.PtrToStructure<ArrowSchema>(child.Dictionary).Format));

                ArrowArray column = ArrowChildArray(array, 0);
                Assert.Equal((20L, 2L, 3L), (column.Length, column.NBuffers, column.NullCount));
                ArrowArray dictionary = Marshal.PtrToStructure<ArrowArray>(column.Dictionary);
                Assert.Equal((Cities.Length + 1L, 3L, 0L), (dictionary.Length, dictionary.NBuffers, dictionary.NullCount));
                byte[] validity = new byte[3];
                Marshal.Copy(ArrowBuffer(column, 0), validity, 0, validity.Length);
                int[] indices = new int[20];
                Marshal.Copy(ArrowBuffer(column, 1), indices, 0, indices.Length);
                int[] offsets = new int[dictionary.Length + 1];
                Marshal.Copy(ArrowBuffer(dictionary, 1), offsets, 0, offsets.Length);
                byte[] data = new byte[offsets[^1]];
                Marshal.Copy(ArrowBuffer(dictionary, 2), data, 0, data.Length);
                for (int i = 0; i < indices.Length; i++)
                {
                    bool valid = (validity[i >> 3] & (1 << (i & 7))) != 0;
                    string? actual = valid ? Encoding.UTF8.GetString(data, offsets[indices[i]], offsets[indices[i] + 1] - offsets[indices[i]]) : null;
                    Assert.Equal(i % 11 == 5 ? "42" : CityAt(i), actual);
                }
            }
            finally
            {
                ExercisedReleaseArrow(ref array, ref schema);
                ReadApi.Close(handle);
            }
            Assert.Equal((IntPtr.Zero, IntPtr.Zero), (array.Release, schema.Release));
        }

        [Fact]
        public unsafe void OpenArrowStream_Should_Emit_A_Dictionary_Per_Batch()
        {
            Assert.Equal(NativeStatus.Ok, ReadApi.OpenMemory(CityXlsb(20, sharedStrings: true), NativeFormat.Xlsb, out NativeHandle? handle));
            NativeColumnSpec[] specs = [new() { Names = ["city"], Type = NativeColumnType.StringDictionary }];
            Assert.Equal(NativeStatus.Ok, ArrowApi.OpenArrowStream(handle, 0, specs, 1, maxRows: 8, out ArrowArrayStream stream));
            try
            {
                ArrowArray batch;
                Assert.Equal(0, ArrowApi.ArrowStreamGetNextCore(&stream, &batch));
                ArrowArray column = ArrowChildArray(batch, 0);
                Assert.Equal(8, column.Length);
                Assert.Equal(Cities.Length + 1, Marshal.PtrToStructure<ArrowArray>(column.Dictionary).Length);
                ArrowApi.ReleaseArrowArray((IntPtr)(&batch));
                Assert.Equal(IntPtr.Zero, batch.Release);
            }
            finally
            {
                ArrowApi.ArrowStreamReleaseCore(&stream);
                ReadApi.Close(handle);
            }
        }

        [Fact]
        public void ParseTyped_Should_Reject_A_Dictionary_Column()
        {
            Assert.Equal(NativeStatus.Ok, ReadApi.OpenMemory(CityXlsb(3, sharedStrings: true), NativeFormat.Xlsb, out NativeHandle? handle));
            try
            {
                NativeColumnSpec[] specs = [new() { Names = ["city"], Type = NativeColumnType.StringDictionary }];
                Assert.Equal(NativeStatus.InvalidArgument, TypedApi.ParseTyped(handle, 0, specs, 1, out _));
                Assert.Contains("Arrow outputs", NativeApi.LastErrorText(), StringComparison.Ordinal);
                Assert.Equal(NativeStatus.InvalidArgument, TypedApi.OpenTypedReader(handle, 0, specs, 1, 0, out _));
            }
            finally
            {
                ReadApi.Close(handle);
            }
        }
    }
}
