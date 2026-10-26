using System.Diagnostics.CodeAnalysis;
using System.Text;
using ExcelReader.Core.Parser;
using ExcelReader.Core.Reader;

namespace ExcelReader.Tests.Parser
{
    public sealed class EnumParsingTests
    {
        private enum Adjacent : long
        {
            Low = 9007199254740992,
            High = 9007199254740993,
        }

        private enum Unsigned : ulong
        {
            Low = 0,
            High = ulong.MaxValue,
        }

        [SuppressMessage("Design", "CA1069:Enums values should not be duplicated", Justification = "The alias is what the test covers.")]
        private enum Aliased
        {
            First = 1,
            Alias = 1,
        }

        private sealed class AliasRow
        {
            public Aliased Value { get; set; }
        }

        private static bool TryParse<TEnum>(CellType type, string text, out TEnum value)
            where TEnum : struct, Enum
        {
            Cell cell = new(type, Encoding.UTF8.GetBytes(text));
            return ExcelCellReaders.Enum(in cell, default, out value);
        }

        [Theory]
        [InlineData("9007199254740992", 9007199254740992L)]
        [InlineData("9007199254740993", 9007199254740993L)]
        public void Should_PickTheExactMember_When_AdjacentValuesShareADouble(string text, long expected)
        {
            Assert.True(TryParse(CellType.Number, text, out Adjacent number));
            Assert.Equal((Adjacent)expected, number);
            Assert.True(TryParse(CellType.ExcelString, text, out Adjacent textual));
            Assert.Equal((Adjacent)expected, textual);
        }

        [Theory]
        [InlineData("9007199254740993.5")]
        [InlineData("1E300")]
        [InlineData("NaN")]
        [InlineData("-1")]
        [InlineData("18446744073709551616")]
        public void Should_Reject_When_TheNumberIsNotADeclaredIntegralValue(string text)
        {
            Assert.False(TryParse(CellType.Number, text, out Unsigned _));
        }

        [Theory]
        [InlineData("Low", 0UL)]
        [InlineData("high", ulong.MaxValue)]
        [InlineData("0", 0UL)]
        [InlineData("18446744073709551615", ulong.MaxValue)]
        public void Should_ParseAnUnsignedEnum_When_AMemberExceedsInt64(string text, ulong expected)
        {
            Assert.True(TryParse(CellType.ExcelString, text, out Unsigned textual));
            Assert.Equal((Unsigned)expected, textual);
            if (char.IsDigit(text[0]))
            {
                Assert.True(TryParse(CellType.Number, text, out Unsigned number));
                Assert.Equal((Unsigned)expected, number);
            }
        }

        [Fact]
        public void Should_ReadABinaryNumberAsItsStoredDouble()
        {
            Cell cell = new(CellType.Number, default, 9007199254740992.0, hasNumber: true, styleIndex: 0);

            Assert.True(ExcelCellReaders.Enum(in cell, default, out Adjacent value));
            Assert.Equal(Adjacent.Low, value);
        }

        [Theory]
        [InlineData("First")]
        [InlineData("alias")]
        [InlineData("1")]
        public void Should_ParseEveryAliasName(string text)
        {
            Assert.True(TryParse(CellType.ExcelString, text, out Aliased value));
            Assert.Equal(Aliased.First, value);
        }

        [Fact]
        public void Should_RejectAnUndefinedValue()
        {
            Assert.False(TryParse(CellType.ExcelString, "Second", out Aliased _));
            Assert.False(TryParse(CellType.Number, "2", out Aliased _));
        }

        [Fact]
        public async Task Should_ParseAnAliasName_When_ReadingThroughAttributes()
        {
            await using var ms = await TypedWorkbook.BuildAsync(["Value"], ["Alias"], ["FIRST"]);
            await using var reader = await Excel.FromXlsxAsync(ms, ct: TestContext.Current.CancellationToken);

            AliasRow[] rows = ExcelParser.FromAttributes<AliasRow>().Parse(reader.FirstSheet).ToArray();

            Assert.Equal([Aliased.First, Aliased.First], rows.Select(static row => row.Value));
        }
    }
}
