using System.Globalization;
using ExcelReader.Core.Parser;
using ExcelReader.Core.Reader;

namespace ExcelReader.Tests
{
    public class ParserFeatureTests
    {
        private enum Status
        {
            Unknown = 0,
            Active = 1,
            Closed = 2,
        }

        private enum LargeStatus : long
        {
            High = 3_000_000_000,
        }

        private sealed class LargeEnumRow
        {
            public LargeStatus? Status { get; set; }
        }

        private sealed class MoneyRow
        {
            public string? Name { get; set; }
            public decimal Amount { get; set; }
        }

        private sealed class TypedRow
        {
            public Status Status { get; set; }
            public Status? OptionalStatus { get; set; }
            public Guid Id { get; set; }
            public Guid? OptionalId { get; set; }
        }


        [Fact]
        public async Task ParseAcceptsAutoDetectedReader()
        {
            await using var ms = await TypedWorkbook.BuildAsync(["Name", "Amount"], ["Alice", 12.5]);
            using IExcelRowReader reader = Excel.Open(ms);

            MoneyRow row = new ExcelParser<MoneyRow>().Parse(reader).Single();

            Assert.Equal("Alice", row.Name);
            Assert.Equal(12.5m, row.Amount);
        }

        [Fact]
        public async Task ParseAsyncAcceptsAutoDetectedReader()
        {
            await using var ms = await TypedWorkbook.BuildAsync(["Name", "Amount"], ["Bob", 7.0]);
            await using IExcelRowReader reader = await Excel.OpenAsync(ms, ct: TestContext.Current.CancellationToken);

            var rows = new List<MoneyRow>();
            await foreach (MoneyRow row in new ExcelParser<MoneyRow>().ParseAsync(reader, TestContext.Current.CancellationToken))
            {
                rows.Add(row);
            }

            MoneyRow only = Assert.Single(rows);
            Assert.Equal("Bob", only.Name);
            Assert.Equal(7.0m, only.Amount);
        }


        [Fact]
        public async Task PtBrCultureParsesCommaDecimalAndThousandsSeparator()
        {
            await using var ms = await TypedWorkbook.BuildAsync(["Name", "Amount"], ["Conta", "1.234,56"]);
            await using var reader = await Excel.FromXlsxAsync(ms, ct: TestContext.Current.CancellationToken);
            var config = new ExcelParserConfig { Culture = CultureInfo.GetCultureInfo("pt-BR") };

            MoneyRow row = new ExcelParser<MoneyRow>(config).Parse(reader).Single();

            Assert.Equal(1234.56m, row.Amount);
        }

        [Fact]
        public async Task InvariantCultureRejectsCommaDecimal()
        {
            await using var ms = await TypedWorkbook.BuildAsync(["Name", "Amount"], ["Conta", "1.234,56"]);
            await using var reader = await Excel.FromXlsxAsync(ms, ct: TestContext.Current.CancellationToken);

            MoneyRow row = new ExcelParser<MoneyRow>().Parse(reader).Single();

            Assert.Equal(0m, row.Amount);
        }


        [Fact]
        public async Task ThrowOnParseFailureThrowsForUnparseableNonRequiredColumn()
        {
            await using var ms = await TypedWorkbook.BuildAsync(["Name", "Amount"], ["Conta", "not-a-number"]);
            await using var reader = await Excel.FromXlsxAsync(ms, ct: TestContext.Current.CancellationToken);
            var config = new ExcelParserConfig { ThrowOnParseFailure = true };

            ExcelParseException ex = Assert.Throws<ExcelParseException>(
                () => new ExcelParser<MoneyRow>(config).Parse(reader).ToList());
            Assert.Equal("Amount", ex.ColumnName);
            Assert.Equal("not-a-number", ex.RawValue);
        }

        [Fact]
        public async Task DefaultConfigLeavesUnparseableNonRequiredColumnAtDefault()
        {
            await using var ms = await TypedWorkbook.BuildAsync(["Name", "Amount"], ["Conta", "not-a-number"]);
            await using var reader = await Excel.FromXlsxAsync(ms, ct: TestContext.Current.CancellationToken);

            MoneyRow row = new ExcelParser<MoneyRow>().Parse(reader).Single();

            Assert.Equal(0m, row.Amount);
        }


        [Fact]
        public async Task EnumColumnsParseByNameAndNumber()
        {
            var id = Guid.NewGuid();
            var optId = Guid.NewGuid();
            await using var ms = await TypedWorkbook.BuildAsync(
                ["Status", "OptionalStatus", "Id", "OptionalId"],
                ["Active", 2, id.ToString(), optId.ToString()]);
            await using var reader = await Excel.FromXlsxAsync(ms, ct: TestContext.Current.CancellationToken);

            TypedRow row = new ExcelParser<TypedRow>().Parse(reader).Single();

            Assert.Equal(Status.Active, row.Status);
            Assert.Equal(Status.Closed, row.OptionalStatus);
            Assert.Equal(id, row.Id);
            Assert.Equal(optId, row.OptionalId);
        }

        [Fact]
        public async Task EnumIsCaseInsensitiveAndInvalidKeepsDefault()
        {
            await using var ms = await TypedWorkbook.BuildAsync(
                ["Status", "OptionalStatus"],
                ["active", "garbage"]);
            await using var reader = await Excel.FromXlsxAsync(ms, ct: TestContext.Current.CancellationToken);

            TypedRow row = new ExcelParser<TypedRow>().Parse(reader).Single();

            Assert.Equal(Status.Active, row.Status);
            Assert.Null(row.OptionalStatus);
        }

        [Fact]
        public async Task EnumFromNumericTextCellBindsByValue()
        {
            await using var ms = await TypedWorkbook.BuildAsync(
                ["Status", "OptionalStatus"],
                ["2", "1"]);
            await using var reader = await Excel.FromXlsxAsync(ms, ct: TestContext.Current.CancellationToken);

            TypedRow row = new ExcelParser<TypedRow>().Parse(reader).Single();

            Assert.Equal(Status.Closed, row.Status);
            Assert.Equal(Status.Active, row.OptionalStatus);
        }

        [Fact]
        public async Task InvalidGuidKeepsDefault()
        {
            await using var ms = await TypedWorkbook.BuildAsync(
                ["Id", "OptionalId"],
                ["not-a-guid", "also-bad"]);
            await using var reader = await Excel.FromXlsxAsync(ms, ct: TestContext.Current.CancellationToken);

            TypedRow row = new ExcelParser<TypedRow>().Parse(reader).Single();

            Assert.Equal(Guid.Empty, row.Id);
            Assert.Null(row.OptionalId);
        }

        [Fact]
        public async Task LongBackedEnumParsesWithoutTruncatingFractionalNumbers()
        {
            await using var ms = await TypedWorkbook.BuildAsync(
                ["Status"],
                [3_000_000_000d],
                [3_000_000_000.5d]);
            await using var reader = await Excel.FromXlsxAsync(ms, ct: TestContext.Current.CancellationToken);

            LargeEnumRow[] rows = new ExcelParser<LargeEnumRow>().Parse(reader).ToArray();

            Assert.Equal(LargeStatus.High, rows[0].Status);
            Assert.Null(rows[1].Status);
        }
    }
}
