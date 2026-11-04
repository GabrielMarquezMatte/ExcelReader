using ExcelReader.Core.Reader;
using ExcelReader.Core.Writer;
using ExcelReader.Core.Writer.Internal;

namespace ExcelReader.Tests.Writer.Styles
{
    public class StyleTableTests
    {
        private static ExcelColor Color(int rgb)
        {
            return ExcelColor.FromRgb(rgb);
        }

        [Fact]
        public void StylesDifferingOnlyInFillShareTheirFont()
        {
            var table = new StyleTable(StyleLimits.Ooxml);
            int a = table.Add(new CellStyle { Bold = true, FillColor = Color(1) });
            int b = table.Add(new CellStyle { Bold = true, FillColor = Color(2) });
            StyleComponents components = table.AssignComponents();
            Assert.Equal(components.FontId(a), components.FontId(b));
            Assert.NotEqual(components.FillId(a), components.FillId(b));
        }

        [Fact]
        public void CustomFillsStartAfterTheTwoRequiredFills()
        {
            var table = new StyleTable(StyleLimits.Ooxml);
            int a = table.Add(new CellStyle { FillColor = Color(1) });
            StyleComponents components = table.AssignComponents();
            Assert.Equal(2, components.FillId(a));
            Assert.Equal(0, components.FillId(0));
        }

        [Fact]
        public void BoldAndItalicMapToSeparateFonts()
        {
            var table = new StyleTable(StyleLimits.Ooxml);
            int bold = table.Add(new CellStyle { Bold = true });
            int italic = table.Add(new CellStyle { Italic = true });
            StyleComponents components = table.AssignComponents();
            Assert.Equal(3, components.Fonts.Count);
            Assert.NotEqual(components.FontId(bold), components.FontId(italic));
            Assert.Equal(0, components.FontId(0));
        }

        [Fact]
        public void FiftySeventhDistinctColorThrowsForBiff8()
        {
            var table = new StyleTable(StyleLimits.Biff8);
            for (int i = 1; i <= 56; i++)
            {
                table.Add(new CellStyle { FillColor = Color(i) });
            }
            var ex = Assert.Throws<ExcelLimitExceededException>(() => table.Add(new CellStyle { FillColor = Color(57) }));
            Assert.Equal("Colors", ex.LimitName);
        }

        [Fact]
        public void ReusingAPaletteColorDoesNotCountAgain()
        {
            var table = new StyleTable(StyleLimits.Biff8);
            for (int i = 1; i <= 56; i++)
            {
                table.Add(new CellStyle { FillColor = Color(i) });
            }
            table.Add(new CellStyle { FontColor = Color(7) });
            Assert.Equal(56, table.PaletteColors.Count);
        }

        [Fact]
        public void OoxmlTableHasNoColorLimit()
        {
            var table = new StyleTable(StyleLimits.Ooxml);
            for (int i = 1; i <= 100; i++)
            {
                table.Add(new CellStyle { FillColor = Color(i) });
            }
            Assert.Equal(100, table.PaletteColors.Count);
        }

        [Fact]
        public void CellStyleLimitThrowsAndLeavesTheTableUsable()
        {
            var table = new StyleTable(new StyleLimits(4, 512, int.MaxValue));
            table.Add(new CellStyle { Bold = true });
            table.Add(new CellStyle { Italic = true });
            var ex = Assert.Throws<ExcelLimitExceededException>(() => table.Add(new CellStyle { Underline = true }));
            Assert.Equal("CellStyles", ex.LimitName);
            Assert.Equal(4, table.Count);
        }

        [Fact]
        public void FontLimitThrowsAndCountsTheReservedBiff8Index()
        {
            var table = new StyleTable(new StyleLimits(65_490, 3, int.MaxValue));
            table.Add(new CellStyle { Bold = true });
            table.Add(new CellStyle { Italic = true });
            var ex = Assert.Throws<ExcelLimitExceededException>(() => table.Add(new CellStyle { Underline = true }));
            Assert.Equal("Fonts", ex.LimitName);

            var biff8 = new StyleTable(new StyleLimits(65_490, 3, int.MaxValue) { ReservedFontIndices = 1 });
            biff8.Add(new CellStyle { Bold = true });
            Assert.Throws<ExcelLimitExceededException>(() => biff8.Add(new CellStyle { Italic = true }));
        }
    }
}
