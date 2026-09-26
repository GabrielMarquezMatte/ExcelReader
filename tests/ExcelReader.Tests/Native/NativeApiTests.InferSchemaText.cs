using ExcelReader.Native;
using ExcelReader.Native.Reading;
using ExcelReader.Native.Typed;

namespace ExcelReader.Tests.Native
{
    public sealed partial class NativeApiTests
    {
        private static string FixtureCsvPath()
        {
            string path = Path.Combine(AppContext.BaseDirectory, "data", "65K_Records_Data.csv");
            if (!File.Exists(path))
            {
                path = Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..",
                    "ExcelReader.Benchmarks", "Data", "65K_Records_Data.csv");
            }
            Assert.True(File.Exists(path), $"CSV fixture not found at {path}");
            return path;
        }

        [Fact]
        public void InferSchema_Should_Type_Text_Cells_When_ParseText_Is_Set()
        {
            Assert.Equal(NativeStatus.Ok, OpenPath(FixtureCsvPath(), NativeFormat.Csv, out NativeHandle? handle));
            try
            {
                Assert.Equal(NativeStatus.Ok, ReadApi.InferSchema(handle, headerRow: 1, sampleSize: 100, ReadApi.InferParseText, out NativeInferredSchema schema));
                (string? Name, int Index, int Type, bool Nullable)[] columns;
                try
                {
                    columns = DecodeSchema(schema);
                }
                finally
                {
                    ReadApi.FreeSchema(ref schema);
                }

                int s = NativeColumnType.String, d = NativeColumnType.Date, i = NativeColumnType.Int64, f = NativeColumnType.Float64;
                Assert.Equal([s, s, s, s, s, d, i, d, i, f, f, f, f, f], Array.ConvertAll(columns, c => c.Type));

                NativeColumnSpec[] specs = Array.ConvertAll(columns, c => new NativeColumnSpec { Names = [c.Name!], Type = c.Type, Nullable = c.Nullable });
                Assert.Equal(NativeStatus.Ok, TypedApi.ParseTyped(handle, specs, headerRow: 1, out NativeTable table));
                try
                {
                    Assert.Equal(65_535, table.RowCount);
                }
                finally
                {
                    TypedApi.FreeTable(ref table);
                }
            }
            finally
            {
                ReadApi.Close(handle);
            }
        }

        [Fact]
        public void InferSchema_Should_Match_The_Flagless_Result_When_Flags_Are_Zero()
        {
            using MemoryStream ms = BuildInferSchemaFixture();
            Assert.Equal(NativeStatus.Ok, ReadApi.OpenMemory(ms.ToArray(), NativeFormat.Xlsx, out NativeHandle? handle));
            try
            {
                Assert.Equal(NativeStatus.Ok, ReadApi.InferSchema(handle, headerRow: 1, sampleSize: 100, out NativeInferredSchema plain));
                Assert.Equal(NativeStatus.Ok, ReadApi.InferSchema(handle, headerRow: 1, sampleSize: 100, 0, out NativeInferredSchema flagged));
                try
                {
                    Assert.Equal(DecodeSchema(plain), DecodeSchema(flagged));
                }
                finally
                {
                    ReadApi.FreeSchema(ref plain);
                    ReadApi.FreeSchema(ref flagged);
                }
            }
            finally
            {
                ReadApi.Close(handle);
            }
        }

        [Fact]
        public void InferSchema_Should_Reject_Unknown_Flags()
        {
            using MemoryStream ms = BuildInferSchemaFixture();
            Assert.Equal(NativeStatus.Ok, ReadApi.OpenMemory(ms.ToArray(), NativeFormat.Xlsx, out NativeHandle? handle));
            try
            {
                Assert.Equal(NativeStatus.InvalidArgument, ReadApi.InferSchema(handle, headerRow: 1, sampleSize: 100, 2, out NativeInferredSchema schema));
                Assert.Equal(IntPtr.Zero, schema.Columns);
                Assert.Contains("flags", NativeApi.LastErrorText(), StringComparison.Ordinal);
            }
            finally
            {
                ReadApi.Close(handle);
            }
        }
    }
}
