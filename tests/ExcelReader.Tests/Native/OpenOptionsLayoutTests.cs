using ExcelReader.Native;
using ExcelReader.Native.Writer;

namespace ExcelReader.Tests.Native
{
    public sealed unsafe class OpenOptionsLayoutTests
    {
        [Fact]
        public void The_6_0_Layout_Ends_Where_The_New_Fields_Begin()
        {
            Assert.Equal(88, NativeOpenOptionsRaw.V6Size);
            Assert.Equal(NativeOpenOptionsRaw.V6Size, (int)System.Runtime.InteropServices.Marshal.OffsetOf<NativeOpenOptionsRaw>(nameof(NativeOpenOptionsRaw.SourceCacheBytes)));
            Assert.Equal(112, sizeof(NativeOpenOptionsRaw));
        }

        [Fact]
        public void The_Checksum_Switch_Sits_In_The_Former_Tail_Padding()
        {
            Assert.Equal(108, (int)System.Runtime.InteropServices.Marshal.OffsetOf<NativeOpenOptionsRaw>(nameof(NativeOpenOptionsRaw.VerifyZipChecksums)));
        }

        [Theory]
        [InlineData(0, true)]
        [InlineData(1, false)]
        [InlineData(2, true)]
        public void The_Checksum_Switch_Reaches_The_Reader_Options(int state, bool expected)
        {
            NativeOpenOptionsRaw raw = new() { StructSize = sizeof(NativeOpenOptionsRaw), VerifyZipChecksums = state };

            Assert.True(NativeOpenOptions.TryDecode(raw, out NativeOpenOptions options, out string? error), error);
            Assert.Equal(expected, options.ToExcelReaderOptions().VerifyZipChecksums);
        }

        [Fact]
        public void An_Unknown_Checksum_State_Is_Rejected()
        {
            NativeOpenOptionsRaw raw = new() { StructSize = sizeof(NativeOpenOptionsRaw), VerifyZipChecksums = 3 };

            Assert.False(NativeOpenOptions.TryDecode(raw, out _, out string? error));
            Assert.Contains("verify_zip_checksums", error, StringComparison.Ordinal);
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
        public void CellStyleLayoutMatchesTheHeader()
        {
            NativeCellStyleRaw style = default;
            byte* start = (byte*)&style;

            Assert.Equal(8, sizeof(NativeBorderRaw));
            Assert.Equal(104, sizeof(NativeCellStyleRaw));
            Assert.Equal(
                [0, 4, 8, 16, 24, 28, 32, 40, 44, 48, 52, 56, 60, 64, 68, 76, 84, 92],
                [
                    (int)((byte*)&style.StructSize - start), (int)((byte*)&style.NumberFormatLen - start),
                    (int)((byte*)&style.NumberFormat - start), (int)((byte*)&style.FontName - start),
                    (int)((byte*)&style.FontNameLen - start), (int)((byte*)&style.FontFlags - start),
                    (int)((byte*)&style.FontSize - start), (int)((byte*)&style.FontColor - start),
                    (int)((byte*)&style.FillColor - start), (int)((byte*)&style.HorizontalAlignment - start),
                    (int)((byte*)&style.VerticalAlignment - start), (int)((byte*)&style.WrapText - start),
                    (int)((byte*)&style.Indent - start), (int)((byte*)&style.Rotation - start),
                    (int)((byte*)&style.Left - start), (int)((byte*)&style.Right - start),
                    (int)((byte*)&style.Top - start), (int)((byte*)&style.Bottom - start),
                ]);
        }

        [Fact]
        public void SheetOptionsLayoutMatchesTheHeader()
        {
            NativeSheetOptionsRaw options = default;
            byte* start = (byte*)&options;

            Assert.Equal(64, sizeof(NativeSheetOptionsRaw));
            Assert.Equal(40, NativeSheetOptionsRaw.V6Size);
            Assert.Equal(
                [0, 4, 8, 16, 24, 32, 40, 48, 56],
                [
                    (int)((byte*)&options.StructSize - start), (int)((byte*)&options.TableNameLen - start),
                    (int)((byte*)&options.TableName - start), (int)((byte*)&options.TableStyleLen - start),
                    (int)((byte*)&options.TableStyle - start), (int)((byte*)&options.TableFlags - start),
                    (int)((byte*)&options.ColumnStyles - start), (int)((byte*)&options.ColumnWidths - start),
                    (int)((byte*)&options.HeaderStyle - start),
                ]);
            Assert.Equal(40, sizeof(NativeWriteOptionsRaw));
        }

        [Fact]
        public void AbiVersionIsSeven()
        {
            Assert.Equal(7, ((delegate* unmanaged<int>)&Exports.AbiVersion)());
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
