using ExcelReader.Native;

namespace ExcelReader.Tests.Native
{
    public sealed unsafe class OpenOptionsLayoutTests
    {
        [Fact]
        public void The_6_0_Layout_Ends_Where_The_New_Fields_Begin()
        {
            Assert.Equal(88, NativeOpenOptionsRaw.V6Size);
            Assert.Equal(112, sizeof(NativeOpenOptionsRaw));
        }

        [Fact]
        public void A_6_0_Sized_Struct_Is_Read_As_A_Prefix()
        {
            NativeOpenOptionsRaw caller = new()
            {
                StructSize = NativeOpenOptionsRaw.V6Size,
                MaxCellBytes = 1234,
                SourceCacheBytes = -99,
                MaxBufferedBytes = -99,
                SourceBlockSize = -99,
            };
            Assert.True(NativeOpenOptionsRaw.TryRead(&caller, out NativeOpenOptionsRaw raw, out string? error), error);
            Assert.Equal(1234, raw.MaxCellBytes);
            Assert.Equal(0, raw.SourceCacheBytes);
            Assert.Equal(0, raw.MaxBufferedBytes);
            Assert.Equal(0, raw.SourceBlockSize);
            Assert.Equal(sizeof(NativeOpenOptionsRaw), raw.StructSize);
        }

        [Fact]
        public void A_Full_Sized_Struct_Is_Read_Whole()
        {
            NativeOpenOptionsRaw caller = new()
            {
                StructSize = sizeof(NativeOpenOptionsRaw),
                SourceCacheBytes = 1 << 20,
                MaxBufferedBytes = 4096,
                SourceBlockSize = -1,
            };
            Assert.True(NativeOpenOptionsRaw.TryRead(&caller, out NativeOpenOptionsRaw raw, out _));
            Assert.Equal(1 << 20, raw.SourceCacheBytes);
            Assert.Equal(4096, raw.MaxBufferedBytes);
            Assert.Equal(-1, raw.SourceBlockSize);
        }

        [Theory]
        [InlineData(0)]
        [InlineData(1)]
        [InlineData(87)]
        [InlineData(113)]
        public void Any_Other_Size_Is_Rejected(int size)
        {
            NativeOpenOptionsRaw caller = new() { StructSize = size };
            Assert.False(NativeOpenOptionsRaw.TryRead(&caller, out _, out string? error));
            Assert.Contains("struct_size", error, StringComparison.Ordinal);
        }

        [Fact]
        public void Decoding_Keeps_A_Negative_Block_Size_And_Rejects_Negative_Byte_Counts()
        {
            NativeOpenOptionsRaw raw = new()
            {
                StructSize = sizeof(NativeOpenOptionsRaw),
                SourceBlockSize = -1,
                SourceCacheBytes = 1 << 20,
                MaxBufferedBytes = 10,
            };
            Assert.True(NativeOpenOptions.TryDecode(raw, out NativeOpenOptions options, out string? error), error);
            Assert.Equal(-1, options.SourceBlockSize);
            Assert.Equal(1 << 20, options.SourceCacheBytes);
            Assert.Equal(10, options.MaxBufferedBytes);

            Assert.False(NativeOpenOptions.TryDecode(raw with { SourceCacheBytes = -1 }, out _, out _));
            Assert.False(NativeOpenOptions.TryDecode(raw with { MaxBufferedBytes = -1 }, out _, out _));
        }

        [Fact]
        public void Zero_Leaves_Every_New_Field_At_Its_Default()
        {
            NativeOpenOptionsRaw raw = new() { StructSize = sizeof(NativeOpenOptionsRaw) };
            Assert.True(NativeOpenOptions.TryDecode(raw, out NativeOpenOptions options, out _));
            Assert.Null(options.SourceBlockSize);
            Assert.Null(options.SourceCacheBytes);
            Assert.Null(options.MaxBufferedBytes);
        }
    }
}
