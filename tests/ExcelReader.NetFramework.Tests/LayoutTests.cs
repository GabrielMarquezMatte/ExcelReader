using System.Runtime.InteropServices;
using ExcelReader.NetFramework.Interop;

namespace ExcelReader.NetFramework.Tests
{
    public sealed class LayoutTests
    {
        private static void Layout<T>(int size, params (string Field, int Offset)[] fields)
        {
            Assert.Equal(size, Marshal.SizeOf<T>());
            foreach (var (field, offset) in fields)
                Assert.Equal(offset, (int)Marshal.OffsetOf<T>(field));
        }

        private static void RequireX64() => Assert.SkipUnless(Environment.Is64BitProcess, "x64 layout");

        [Fact]
        public void XlRowCell_Layout()
        {
            RequireX64();
            Layout<XlRowCell>(24, (nameof(XlRowCell.Value), 16));
        }

        [Fact]
        public void XlRow_Layout()
        {
            RequireX64();
            Layout<XlRow>(16, (nameof(XlRow.Cells), 8));
        }

        [Fact]
        public void XlColumnSpec_Layout()
        {
            RequireX64();
            Layout<XlColumnSpec>(32,
                (nameof(XlColumnSpec.Names), 0), (nameof(XlColumnSpec.NameLens), 8), (nameof(XlColumnSpec.NameCount), 16),
                (nameof(XlColumnSpec.Index), 20), (nameof(XlColumnSpec.Type), 24), (nameof(XlColumnSpec.Nullable), 28));
        }

        [Fact]
        public void XlColumn_Layout()
        {
            RequireX64();
            Layout<XlColumn>(48,
                (nameof(XlColumn.Length), 8), (nameof(XlColumn.Values), 16), (nameof(XlColumn.Validity), 24),
                (nameof(XlColumn.Data), 32), (nameof(XlColumn.DataLen), 40));
        }

        [Fact]
        public void XlTable_Layout()
        {
            RequireX64();
            Layout<XlTable>(24, (nameof(XlTable.RowCount), 8), (nameof(XlTable.Columns), 16));
        }

        [Fact]
        public void XlInferredSchema_Layout()
        {
            RequireX64();
            Layout<XlInferredSchema>(16, (nameof(XlInferredSchema.ColumnCount), 8));
        }

        [Fact]
        public void XlOpenOptions_Layout()
        {
            RequireX64();
            Layout<XlOpenOptions>(112,
                (nameof(XlOpenOptions.SourceCacheBytes), 88), (nameof(XlOpenOptions.MaxBufferedBytes), 96),
                (nameof(XlOpenOptions.SourceBlockSize), 104));
        }

        [Fact]
        public void XlStream_Layout()
        {
            RequireX64();
            Layout<XlStream>(32, (nameof(XlStream.UserData), 8), (nameof(XlStream.Read), 16), (nameof(XlStream.Release), 24));
        }

        [Fact]
        public void XlTableInfo_Layout() => Assert.Equal(28, Marshal.SizeOf<XlTableInfo>());

        [Fact]
        public void X86_PointerSizedFields_Shrink()
        {
            Assert.SkipWhen(Environment.Is64BitProcess, "x86 layout");
            Assert.Equal(24, Marshal.SizeOf<XlColumnSpec>());
            Assert.Equal(16, Marshal.SizeOf<XlStream>());
        }
    }
}
