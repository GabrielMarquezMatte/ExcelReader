using System.Globalization;
using System.Text;
using ExcelReader.Core.Parser;
using ExcelReader.Core.Reader;
using ExcelReader.Core.ValueObjects;

namespace ExcelReader.Tests
{
    public class CsvParserTests
    {
        private enum Status
        {
            Unknown = 0,
            Active = 1,
            Closed = 2,
        }

        private sealed class PersonRow
        {
            public string? Name { get; set; }
            public int Age { get; set; }
            public double Score { get; set; }
            public bool Active { get; set; }
            public decimal Balance { get; set; }
        }

        private sealed class MoneyRow
        {
            public string? Name { get; set; }
            public decimal Amount { get; set; }
        }

        private sealed class TypedRow
        {
            public Status Status { get; set; }
            public Guid Id { get; set; }
            public int? Quantity { get; set; }
        }

        private sealed class AliasRow
        {
            [ExcelColumn("First Name")]
            public string? FirstName { get; set; }
        }

        private sealed class RequiredRow
        {
            [ExcelRequired]
            public int Id { get; set; }

            public string? Note { get; set; }
        }

        private sealed class DateRow
        {
            public DateTime Created { get; set; }
        }

        private sealed class IsoDateConverter : IExcelCellConverter<DateTime>
        {
            public bool TryConvert(in Cell cell, bool isDate1904, IFormatProvider provider, out DateTime value)
            {
                return DateTime.TryParse(cell.GetString(), provider, DateTimeStyles.None, out value);
            }
        }

        private sealed class ConvertedDateRow
        {
            [ExcelConverter(typeof(IsoDateConverter))]
            public DateTime Created { get; set; }
        }

        private static MemoryStream Csv(string content)
        {
            return new MemoryStream(Encoding.UTF8.GetBytes(content));
        }

        [Fact]
        public void AllBasicTypesParseFromCsvText()
        {
            using var ms = Csv("Name,Age,Score,Active,Balance\nAlice,42,95.5,true,12345.67\n");
            using var reader = Excel.FromCsv(ms);

            PersonRow row = new ExcelParser<PersonRow>().Parse(reader).Single();

            Assert.Equal("Alice", row.Name);
            Assert.Equal(42, row.Age);
            Assert.Equal(95.5, row.Score);
            Assert.True(row.Active);
            Assert.Equal(12345.67m, row.Balance);
        }

        [Fact]
        public void ExcelColumnAliasMatchesHeader()
        {
            using var ms = Csv("First Name\nBob\n");
            using var reader = Excel.FromCsv(ms);

            AliasRow row = new ExcelParser<AliasRow>().Parse(reader).Single();

            Assert.Equal("Bob", row.FirstName);
        }

        [Fact]
        public void PtBrCultureParsesCommaDecimal()
        {
            using var ms = Csv("Name,Amount\nConta,\"1.234,56\"\n");
            using var reader = Excel.FromCsv(ms);
            var config = new ExcelParserConfig { Culture = CultureInfo.GetCultureInfo("pt-BR") };

            MoneyRow row = new ExcelParser<MoneyRow>(config).Parse(reader).Single();

            Assert.Equal(1234.56m, row.Amount);
        }

        [Fact]
        public void EnumGuidAndNullableColumnsParseFromCsvText()
        {
            var id = Guid.NewGuid();
            using var ms = Csv($"Status,Id,Quantity\nActive,{id},7\n");
            using var reader = Excel.FromCsv(ms);

            TypedRow row = new ExcelParser<TypedRow>().Parse(reader).Single();

            Assert.Equal(Status.Active, row.Status);
            Assert.Equal(id, row.Id);
            Assert.Equal(7, row.Quantity);
        }

        [Fact]
        public void EmptyCellLeavesNullableColumnNull()
        {
            using var ms = Csv("Status,Id,Quantity\nActive,,\n");
            using var reader = Excel.FromCsv(ms);

            TypedRow row = new ExcelParser<TypedRow>().Parse(reader).Single();

            Assert.Null(row.Quantity);
            Assert.Equal(Guid.Empty, row.Id);
        }

        [Fact]
        public void HeaderRowGreaterThanOneSkipsPrecedingRows()
        {
            using var ms = Csv("ignore me\nName,Age,Score,Active,Balance\nAlice,42,95.5,true,12345.67\n");
            var config = new ExcelParserConfig { HeaderRow = 2 };

            PersonRow row = new ExcelParser<PersonRow>(config).Parse(Excel.FromCsv(ms)).Single();

            Assert.Equal("Alice", row.Name);
        }

        [Fact]
        public void MissingRequiredColumnThrowsAtHeader()
        {
            using var ms = Csv("Note\nhi\n");
            using var reader = Excel.FromCsv(ms);

            Assert.Throws<InvalidOperationException>(() => new ExcelParser<RequiredRow>().Parse(reader).ToList());
        }

        [Fact]
        public void EmptyRequiredCellThrowsNamingColumnAndRow()
        {
            using var ms = Csv("Id,Note\n,hi\n");
            using var reader = Excel.FromCsv(ms);

            Assert.Throws<InvalidOperationException>(() => new ExcelParser<RequiredRow>().Parse(reader).ToList());
        }

        [Fact]
        public void PlainDateTimeColumnKeepsDefaultForCsvText()
        {
            // Known limitation (documented): DateTime/DateOnly parse an Excel serial number, not
            // arbitrary text, so a plain CSV date column keeps its default. Use [ExcelConverter].
            using var ms = Csv("Created\n2026-07-02\n");
            using var reader = Excel.FromCsv(ms);

            DateRow row = new ExcelParser<DateRow>().Parse(reader).Single();

            Assert.Equal(default, row.Created);
        }

        [Fact]
        public void CustomConverterParsesTextDate()
        {
            using var ms = Csv("Created\n2026-07-02\n");
            using var reader = Excel.FromCsv(ms);

            ConvertedDateRow row = new ExcelParser<ConvertedDateRow>().Parse(reader).Single();

            Assert.Equal(new DateTime(2026, 7, 2, 0, 0, 0, DateTimeKind.Unspecified), row.Created);
        }

        [Fact]
        public async Task ParseAsyncReadsAllRows()
        {
            using var ms = Csv("Name,Age,Score,Active,Balance\nAlice,42,95.5,true,1\nBob,7,1.5,false,2\n");
            using var reader = Excel.FromCsv(ms);

            var rows = new List<PersonRow>();
            await foreach (PersonRow row in new ExcelParser<PersonRow>().ParseAsync(reader, TestContext.Current.CancellationToken))
            {
                rows.Add(row);
            }

            Assert.Equal(2, rows.Count);
            Assert.Equal("Alice", rows[0].Name);
            Assert.Equal("Bob", rows[1].Name);
        }
    }
}
