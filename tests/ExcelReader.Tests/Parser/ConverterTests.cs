using System.Globalization;
using ExcelReader.Core.Parser;
using ExcelReader.Core.Reader;

namespace ExcelReader.Tests.Parser
{
    public class ConverterTests
    {
        public readonly record struct Percent(double Fraction);

        private sealed class BrlMoneyConverter : IExcelCellConverter<decimal>
        {
            public bool TryConvert(in Cell cell, bool isDate1904, IFormatProvider provider, out decimal value)
            {
                string text = cell.GetString().Replace("R$", string.Empty, StringComparison.Ordinal).Trim();
                return decimal.TryParse(text, NumberStyles.Currency, CultureInfo.GetCultureInfo("pt-BR"), out value);
            }
        }

        private sealed class PercentConverter : IExcelCellConverter<Percent>
        {
            public bool TryConvert(in Cell cell, bool isDate1904, IFormatProvider provider, out Percent value)
            {
                string text = cell.GetString().TrimEnd('%');
                if (double.TryParse(text, NumberStyles.Any, provider, out double pct))
                {
                    value = new Percent(pct / 100.0);
                    return true;
                }
                value = default;
                return false;
            }
        }

        private sealed class NullablePercentConverter : IExcelCellConverter<Percent?>
        {
            public bool TryConvert(in Cell cell, bool isDate1904, IFormatProvider provider, out Percent? value)
            {
                string text = cell.GetString().TrimEnd('%');
                if (double.TryParse(text, NumberStyles.Any, provider, out double pct))
                {
                    value = new Percent(pct / 100.0);
                    return true;
                }
                value = null;
                return false;
            }
        }

        private sealed class YearConverter : IExcelCellConverter<int>
        {
            public bool TryConvert(in Cell cell, bool isDate1904, IFormatProvider provider, out int value)
            {
                if (cell.TryGetDateTime(isDate1904, out DateTime dt))
                {
                    value = dt.Year;
                    return true;
                }
                value = 0;
                return false;
            }
        }

        private sealed class InvoiceRow
        {
            [ExcelConverter(typeof(BrlMoneyConverter))]
            public decimal Total { get; set; }

            [ExcelConverter(typeof(PercentConverter))]
            public Percent Tax { get; set; }

            public string? Ref { get; set; }
        }

        private sealed class NullableRow
        {
            [ExcelConverter(typeof(NullablePercentConverter))]
            public Percent? Tax { get; set; }
        }

        private sealed class DatedRow
        {
            [ExcelConverter(typeof(YearConverter))]
            public int Year { get; set; }
        }

        private sealed class MismatchedConverter : IExcelCellConverter<int>
        {
            public bool TryConvert(in Cell cell, bool isDate1904, IFormatProvider provider, out int value)
            {
                value = 0;
                return false;
            }
        }

        private sealed class BadRow
        {
            [ExcelConverter(typeof(MismatchedConverter))]
            public string? Name { get; set; }
        }

        [Fact]
        public async Task ConvertersParseDomainValues()
        {
            await using var ms = await TypedWorkbook.BuildAsync(
                ["Total", "Tax", "Ref"],
                ["R$ 1.234,56", "12.5%", "INV-1"]);
            await using var reader = await Excel.FromXlsxAsync(ms, ct: TestContext.Current.CancellationToken);

            InvoiceRow row = ExcelParser.FromAttributes<InvoiceRow>().Parse(reader).Single();

            Assert.Equal(1234.56m, row.Total);
            Assert.Equal(0.125, row.Tax.Fraction, precision: 10);
            Assert.Equal("INV-1", row.Ref);
        }

        [Fact]
        public async Task ConverterFailureKeepsDefault()
        {
            await using var ms = await TypedWorkbook.BuildAsync(
                ["Total", "Tax", "Ref"],
                ["garbage", "nope", "INV-2"]);
            await using var reader = await Excel.FromXlsxAsync(ms, ct: TestContext.Current.CancellationToken);

            InvoiceRow row = ExcelParser.FromAttributes<InvoiceRow>().Parse(reader).Single();

            Assert.Equal(0m, row.Total);
            Assert.Equal(default, row.Tax);
            Assert.Equal("INV-2", row.Ref);
        }

        [Fact]
        public async Task NullableConverterSetsAndLeavesNull()
        {
            await using var ms = await TypedWorkbook.BuildMultiSheetAsync(
                ("S1", [["Tax"], ["33%"], ["bad"], [null]]));
            await using var reader = await Excel.FromXlsxAsync(ms, ct: TestContext.Current.CancellationToken);

            List<NullableRow> rows = [.. ExcelParser.FromAttributes<NullableRow>().Parse(reader)];

            Assert.Equal(3, rows.Count);
            Assert.Equal(0.33, rows[0].Tax!.Value.Fraction, precision: 10);
            Assert.Null(rows[1].Tax);
            Assert.Null(rows[2].Tax);
        }

        [Fact]
        public void ConverterReceivesDate1904Flag()
        {
            const string styles =
                """<styleSheet><cellXfs count="2"><xf numFmtId="0"/><xf numFmtId="14"/></cellXfs></styleSheet>""";
            using var ms = WorkbookBuilder.Build(
                """<row r="1"><c r="A1" t="inlineStr"><is><t>Year</t></is></c></row><row r="2"><c r="A2" s="1"><v>0</v></c></row>""",
                styles: styles,
                date1904: true);
            using var reader = Excel.FromXlsx(ms);
            Assert.True(reader.IsDate1904);

            DatedRow row = ExcelParser.FromAttributes<DatedRow>().Parse(reader).Single();

            Assert.Equal(1904, row.Year);
        }

        [Fact]
        public async Task MismatchedConverterTypeThrows()
        {
            await using var ms = await TypedWorkbook.BuildAsync(["Name"], ["x"]);
            await using var reader = await Excel.FromXlsxAsync(ms, ct: TestContext.Current.CancellationToken);

            InvalidOperationException ex = Assert.Throws<InvalidOperationException>(
                () => ExcelParser.FromAttributes<BadRow>().Parse(reader).ToList());
            Assert.Contains("IExcelCellConverter", ex.Message, StringComparison.Ordinal);
        }
    }
}
