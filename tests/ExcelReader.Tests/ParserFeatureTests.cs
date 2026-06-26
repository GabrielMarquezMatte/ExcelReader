using System.Globalization;
using ExcelReader.Core.Parser;
using ExcelReader.Core.Reader;

namespace ExcelReader.Tests
{
    // Covers the parser additions: format-agnostic Parse(IExcelRowReader), configurable Culture,
    // and enum/Guid column support.
    public class ParserFeatureTests
    {
        private enum Status
        {
            Unknown = 0,
            Active = 1,
            Closed = 2,
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

        // --- #1 Parse(IExcelRowReader) ---

        [Fact]
        public async Task ParseAcceptsAutoDetectedReader()
        {
            await using var ms = await TypedWorkbook.BuildAsync(["Name", "Amount"], ["Alice", 12.5]);
            // Excel.Open returns the format-agnostic IExcelRowReader.
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

        // --- #2 Culture ---

        [Fact]
        public async Task PtBrCultureParsesCommaDecimalAndThousandsSeparator()
        {
            // Brazilian text cell: "1.234,56" → 1234.56. Inline strings keep the text verbatim.
            await using var ms = await TypedWorkbook.BuildAsync(["Name", "Amount"], ["Conta", "1.234,56"]);
            await using var reader = await Excel.FromAsync(ms, ct: TestContext.Current.CancellationToken);
            var config = new ExcelParserConfig { Culture = CultureInfo.GetCultureInfo("pt-BR") };

            MoneyRow row = new ExcelParser<MoneyRow>(config).Parse(reader).Single();

            Assert.Equal(1234.56m, row.Amount);
        }

        [Fact]
        public async Task InvariantCultureRejectsCommaDecimal()
        {
            // With the default invariant culture, "1.234,56" is not a valid decimal → keeps default.
            await using var ms = await TypedWorkbook.BuildAsync(["Name", "Amount"], ["Conta", "1.234,56"]);
            await using var reader = await Excel.FromAsync(ms, ct: TestContext.Current.CancellationToken);

            MoneyRow row = new ExcelParser<MoneyRow>().Parse(reader).Single();

            Assert.Equal(0m, row.Amount);
        }

        // --- #4 Enum + Guid ---

        [Fact]
        public async Task EnumColumnsParseByNameAndNumber()
        {
            var id = Guid.NewGuid();
            var optId = Guid.NewGuid();
            await using var ms = await TypedWorkbook.BuildAsync(
                ["Status", "OptionalStatus", "Id", "OptionalId"],
                ["Active", 2, id.ToString(), optId.ToString()]); // name, numeric, guid, guid
            await using var reader = await Excel.FromAsync(ms, ct: TestContext.Current.CancellationToken);

            TypedRow row = new ExcelParser<TypedRow>().Parse(reader).Single();

            Assert.Equal(Status.Active, row.Status);          // by name
            Assert.Equal(Status.Closed, row.OptionalStatus);  // by underlying number
            Assert.Equal(id, row.Id);
            Assert.Equal(optId, row.OptionalId);
        }

        [Fact]
        public async Task EnumIsCaseInsensitiveAndInvalidKeepsDefault()
        {
            await using var ms = await TypedWorkbook.BuildAsync(
                ["Status", "OptionalStatus"],
                ["active", "garbage"]); // lowercase name; unparseable
            await using var reader = await Excel.FromAsync(ms, ct: TestContext.Current.CancellationToken);

            TypedRow row = new ExcelParser<TypedRow>().Parse(reader).Single();

            Assert.Equal(Status.Active, row.Status);     // case-insensitive
            Assert.Null(row.OptionalStatus);             // invalid → left null
        }

        [Fact]
        public async Task InvalidGuidKeepsDefault()
        {
            await using var ms = await TypedWorkbook.BuildAsync(
                ["Id", "OptionalId"],
                ["not-a-guid", "also-bad"]);
            await using var reader = await Excel.FromAsync(ms, ct: TestContext.Current.CancellationToken);

            TypedRow row = new ExcelParser<TypedRow>().Parse(reader).Single();

            Assert.Equal(Guid.Empty, row.Id);
            Assert.Null(row.OptionalId);
        }
    }
}
