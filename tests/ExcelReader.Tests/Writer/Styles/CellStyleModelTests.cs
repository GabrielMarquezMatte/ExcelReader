using ExcelReader.Core.Writer;
using ExcelReader.Core.Writer.Xls;
using ExcelReader.Core.Writer.Xlsb;
using ExcelReader.Core.Writer.Xlsx;

namespace ExcelReader.Tests.Writer.Styles
{
    public class CellStyleModelTests
    {
        public static TheoryData<string> Formats => ["xlsx", "xlsb", "xls"];

        private static int AddStyle(string format, CellStyle style)
        {
            using var ms = new MemoryStream();
            if (string.Equals(format, "xlsx", StringComparison.Ordinal))
            {
                using var xlsx = XlsxWorkbookWriter.Create(ms, leaveOpen: true);
                return xlsx.AddStyle(style);
            }
            if (string.Equals(format, "xlsb", StringComparison.Ordinal))
            {
                using var xlsb = XlsbWorkbookWriter.Create(ms, leaveOpen: true);
                return xlsb.AddStyle(style);
            }
            using var xls = XlsWorkbookWriter.Create(ms, leaveOpen: true);
            return xls.AddStyle(style);
        }

        private static void AssertRejected(string format, CellStyle style)
        {
            var ex = Assert.Throws<ArgumentException>(() => AddStyle(format, style));
            Assert.Equal("style", ex.ParamName);
        }

        [Theory]
        [MemberData(nameof(Formats))]
        public void DefaultStyleStaysTheGeneralStyle(string format)
        {
            Assert.Equal(0, AddStyle(format, default));
            Assert.Equal(0, AddStyle(format, new CellStyle()));
        }

        [Fact]
        public void EqualStylesWithBordersShareOneId()
        {
            CellStyle Make() => new() { Left = new CellBorder(ExcelBorderStyle.Thin, ExcelColor.FromRgb(0xFF0000)) };
            using var ms = new MemoryStream();
            using var wb = XlsxWorkbookWriter.Create(ms, leaveOpen: true);
            Assert.Equal(wb.AddStyle(Make()), wb.AddStyle(Make()));
        }

        [Theory]
        [InlineData("xlsx", 0.5)]
        [InlineData("xlsx", 410.0)]
        [InlineData("xlsb", 0.5)]
        [InlineData("xlsb", 410.0)]
        [InlineData("xls", 0.5)]
        [InlineData("xls", 410.0)]
        public void FontSizeOutOfRangeThrows(string format, double size)
            => AssertRejected(format, new CellStyle { FontSize = size });

        [Theory]
        [InlineData("xlsx", 16)]
        [InlineData("xlsx", -1)]
        [InlineData("xlsb", 16)]
        [InlineData("xlsb", -1)]
        [InlineData("xls", 16)]
        [InlineData("xls", -1)]
        public void IndentOutOfRangeThrows(string format, int indent)
            => AssertRejected(format, new CellStyle { Indent = indent });

        [Theory]
        [InlineData("xlsx", 91)]
        [InlineData("xlsx", -91)]
        [InlineData("xlsx", 254)]
        [InlineData("xlsb", 91)]
        [InlineData("xlsb", -91)]
        [InlineData("xlsb", 254)]
        [InlineData("xls", 91)]
        [InlineData("xls", -91)]
        [InlineData("xls", 254)]
        public void RotationOutOfRangeThrows(string format, int rotation)
            => AssertRejected(format, new CellStyle { Rotation = rotation });

        [Theory]
        [MemberData(nameof(Formats))]
        public void StackedTextRotationIsAccepted(string format)
            => Assert.True(AddStyle(format, new CellStyle { Rotation = CellStyle.StackedTextRotation }) > 1);

        [Theory]
        [MemberData(nameof(Formats))]
        public void FontNameLongerThan31CharactersThrows(string format)
            => AssertRejected(format, new CellStyle { FontName = new string('a', 32) });

        [Theory]
        [MemberData(nameof(Formats))]
        public void EmptyOrControlCharacterFontNameThrows(string format)
        {
            AssertRejected(format, new CellStyle { FontName = "" });
            AssertRejected(format, new CellStyle { FontName = "Ar\u0001ial" });
        }

        [Theory]
        [MemberData(nameof(Formats))]
        public void OverlongNumberFormatStillThrows(string format)
            => AssertRejected(format, new CellStyle { NumberFormat = new string('0', 256) });

        [Fact]
        public void FromRgbRejectsValuesOutsideTwentyFourBits()
        {
            Assert.Throws<ArgumentOutOfRangeException>(() => ExcelColor.FromRgb(0x1000000));
            Assert.Throws<ArgumentOutOfRangeException>(() => ExcelColor.FromRgb(-1));
        }
    }
}
