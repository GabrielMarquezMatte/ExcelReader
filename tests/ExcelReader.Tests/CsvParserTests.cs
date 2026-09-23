using System.Collections;
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

        private sealed class MultiAliasRow
        {
            [ExcelColumn("Preferred Name")]
            [ExcelColumn("Legacy Name")]
            public string? Name { get; set; }
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
            public DateOnly? Day { get; set; }
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

            PersonRow row = ExcelParser.FromAttributes<PersonRow>().Parse(reader).Single();

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

            AliasRow row = ExcelParser.FromAttributes<AliasRow>().Parse(reader).Single();

            Assert.Equal("Bob", row.FirstName);
        }

        [Fact]
        public void HigherPriorityAliasReplacesEarlierLowerPriorityBinding()
        {
            using var ms = Csv("Legacy Name,Preferred Name\nOld,New\n");
            using var reader = Excel.FromCsv(ms);

            MultiAliasRow row = ExcelParser.FromAttributes<MultiAliasRow>().Parse(reader).Single();

            Assert.Equal("New", row.Name);
        }

        [Fact]
        public void PtBrCultureParsesCommaDecimal()
        {
            using var ms = Csv("Name,Amount\nConta,\"1.234,56\"\n");
            using var reader = Excel.FromCsv(ms);
            var config = new ExcelParserConfig { Culture = CultureInfo.GetCultureInfo("pt-BR") };

            MoneyRow row = ExcelParser.FromAttributes<MoneyRow>(config).Parse(reader).Single();

            Assert.Equal(1234.56m, row.Amount);
        }

        [Fact]
        public void EnumGuidAndNullableColumnsParseFromCsvText()
        {
            var id = Guid.NewGuid();
            using var ms = Csv($"Status,Id,Quantity\nActive,{id},7\n");
            using var reader = Excel.FromCsv(ms);

            TypedRow row = ExcelParser.FromAttributes<TypedRow>().Parse(reader).Single();

            Assert.Equal(Status.Active, row.Status);
            Assert.Equal(id, row.Id);
            Assert.Equal(7, row.Quantity);
        }

        [Fact]
        public void EnumFromNumericTextBindsByValueInCsv()
        {
            var id = Guid.NewGuid();
            using var ms = Csv($"Status,Id,Quantity\n2,{id},7\n");
            using var reader = Excel.FromCsv(ms);

            TypedRow row = ExcelParser.FromAttributes<TypedRow>().Parse(reader).Single();

            Assert.Equal(Status.Closed, row.Status);
        }

        [Fact]
        public void EmptyCellLeavesNullableColumnNull()
        {
            using var ms = Csv("Status,Id,Quantity\nActive,,\n");
            using var reader = Excel.FromCsv(ms);

            TypedRow row = ExcelParser.FromAttributes<TypedRow>().Parse(reader).Single();

            Assert.Null(row.Quantity);
            Assert.Equal(Guid.Empty, row.Id);
        }

        [Fact]
        public void HeaderRowGreaterThanOneSkipsPrecedingRows()
        {
            using var ms = Csv("ignore me\nName,Age,Score,Active,Balance\nAlice,42,95.5,true,12345.67\n");
            var config = new ExcelParserConfig { HeaderRow = 2 };

            PersonRow row = ExcelParser.FromAttributes<PersonRow>(config).Parse(Excel.FromCsv(ms)).Single();

            Assert.Equal("Alice", row.Name);
        }

        [Fact]
        public void MissingRequiredColumnThrowsAtHeader()
        {
            using var ms = Csv("Note\nhi\n");
            using var reader = Excel.FromCsv(ms);

            Assert.Throws<ExcelParseException>(() => ExcelParser.FromAttributes<RequiredRow>().Parse(reader).ToList());
        }

        [Fact]
        public void EmptyRequiredCellThrowsNamingColumnAndRow()
        {
            using var ms = Csv("Id,Note\n,hi\n");
            using var reader = Excel.FromCsv(ms);

            Assert.Throws<ExcelParseException>(() => ExcelParser.FromAttributes<RequiredRow>().Parse(reader).ToList());
        }

        [Fact]
        public void RequiredCellWithUnparseableValueThrowsAsIfMissing()
        {
            using var ms = Csv("Id,Note\nabc,hi\n");
            using var reader = Excel.FromCsv(ms);

            ExcelParseException ex = Assert.Throws<ExcelParseException>(
                () => ExcelParser.FromAttributes<RequiredRow>().Parse(reader).ToList());
            Assert.Contains("Id", ex.Message, StringComparison.Ordinal);
        }

        [Fact]
        public void ThrowOnParseFailureThrowsExcelParseExceptionForUnparseableColumn()
        {
            using var ms = Csv("Name,Amount\nConta,not-a-number\n");
            using var reader = Excel.FromCsv(ms);
            var config = new ExcelParserConfig { ThrowOnParseFailure = true };

            ExcelParseException ex = Assert.Throws<ExcelParseException>(
                () => ExcelParser.FromAttributes<MoneyRow>(config).Parse(reader).ToList());
            Assert.Equal("Amount", ex.ColumnName);
            Assert.Equal("not-a-number", ex.RawValue);
        }

        private sealed class ShiftRow
        {
            public TimeOnly Start { get; set; }
        }

        private sealed class StampRow
        {
            public DateTime At { get; set; }
            public DateOnly Day { get; set; }
        }

        [Theory]
        [InlineData("2024-01-15")]
        [InlineData("2024-01-15 13:30:05")]
        [InlineData("2024-01-15T13:30:05")]
        [InlineData("2024-01-15T13:30:05.1234567")]
        [InlineData("2024-01-15 13:30:05.1234567")]
        [InlineData("2024-01-15T13:30:05.5")]
        [InlineData("2024-01-15T13:30:05Z")]
        [InlineData("2024-01-15T13:30:05+03:00")]
        [InlineData("01/15/2024")]
        public void TextDatesBindExactlyAsTheCultureParserWould(string text)
        {
            using var ms = Csv($"At,Day\n{text},{text[..Math.Min(10, text.Length)]}\n");
            using var reader = Excel.FromCsv(ms);

            StampRow row = Assert.Single(ExcelParser.FromAttributes<StampRow>().Parse(reader).ToList());

            Assert.Equal(DateTime.Parse(text, CultureInfo.InvariantCulture, DateTimeStyles.None), row.At);
            Assert.Equal(DateOnly.Parse(text[..Math.Min(10, text.Length)], CultureInfo.InvariantCulture, DateTimeStyles.None), row.Day);
        }

        [Theory]
        [InlineData("NaN")]
        [InlineData("Infinity")]
        [InlineData("-Infinity")]
        public void ANonFiniteTimeOnlyCellFailsToParseInsteadOfBecomingMidnight(string text)
        {
            using var ms = Csv($"Start\n{text}\n");
            using var reader = Excel.FromCsv(ms);
            var config = new ExcelParserConfig { ThrowOnParseFailure = true };

            ExcelParseException ex = Assert.Throws<ExcelParseException>(
                () => ExcelParser.FromAttributes<ShiftRow>(config).Parse(reader).ToList());
            Assert.Equal("Start", ex.ColumnName);
        }

        [Fact]
        public void TerminalBlankLineDoesNotYieldPhantomModelOrRequiredFailure()
        {
            using var ms = Csv("Id,Note\n7,valid\n\n");
            using var reader = Excel.FromCsv(ms);

            RequiredRow row = Assert.Single(ExcelParser.FromAttributes<RequiredRow>().Parse(reader).ToList());

            Assert.Equal(7, row.Id);
            Assert.Equal("valid", row.Note);
        }

        [Fact]
        public void PlainDateTimeAndDateOnlyColumnsParseTextNatively()
        {
            using var ms = Csv("Created,Day\n2026-07-02T08:30:00,2026-07-02\n");
            using var reader = Excel.FromCsv(ms);

            DateRow row = ExcelParser.FromAttributes<DateRow>().Parse(reader).Single();

            Assert.Equal(new DateTime(2026, 7, 2, 8, 30, 0, DateTimeKind.Unspecified), row.Created);
            Assert.Equal(new DateOnly(2026, 7, 2), row.Day);
        }

        [Fact]
        public void UnparseableDateColumnKeepsDefault()
        {
            using var ms = Csv("Created,Day\nnot-a-date,\n");
            using var reader = Excel.FromCsv(ms);

            DateRow row = ExcelParser.FromAttributes<DateRow>().Parse(reader).Single();

            Assert.Equal(default, row.Created);
            Assert.Null(row.Day);
        }

        [Fact]
        public void DateColumnRespectsCulture()
        {
            using var ms = Csv("Created,Day\n02/07/2026,02/07/2026\n");
            using var reader = Excel.FromCsv(ms);
            var config = new ExcelParserConfig { Culture = CultureInfo.GetCultureInfo("pt-BR") };

            DateRow row = ExcelParser.FromAttributes<DateRow>(config).Parse(reader).Single();

            Assert.Equal(new DateTime(2026, 7, 2, 0, 0, 0, DateTimeKind.Unspecified), row.Created);
            Assert.Equal(new DateOnly(2026, 7, 2), row.Day);
        }

        [Fact]
        public void CustomConverterStillOverridesDateParsing()
        {
            using var ms = Csv("Created\n2026-07-02\n");
            using var reader = Excel.FromCsv(ms);

            ConvertedDateRow row = ExcelParser.FromAttributes<ConvertedDateRow>().Parse(reader).Single();

            Assert.Equal(new DateTime(2026, 7, 2, 0, 0, 0, DateTimeKind.Unspecified), row.Created);
        }

        [Fact]
        public async Task ParseAsyncReadsAllRows()
        {
            using var ms = Csv("Name,Age,Score,Active,Balance\nAlice,42,95.5,true,1\nBob,7,1.5,false,2\n");
            using var reader = Excel.FromCsv(ms);

            var rows = new List<PersonRow>();
            await foreach (PersonRow row in ExcelParser.FromAttributes<PersonRow>().Parse(reader).WithCancellation(TestContext.Current.CancellationToken))
            {
                rows.Add(row);
            }

            Assert.Equal(2, rows.Count);
            Assert.Equal("Alice", rows[0].Name);
            Assert.Equal("Bob", rows[1].Name);
        }

        [Fact]
        public void NonGenericEnumerableGetEnumeratorWorks()
        {
            using var ms = Csv("Name\nAlice\n");
            using var reader = Excel.FromCsv(ms);
            IEnumerable enumerable = ExcelParser.FromAttributes<PersonRow>().Parse(reader);

            IEnumerator e = enumerable.GetEnumerator();

            Assert.True(e.MoveNext());
            Assert.Throws<NotSupportedException>(() => e.Current);
        }

        [Fact]
        public void EnumeratorResetThrows()
        {
            using var ms = Csv("Name\nAlice\n");
            using var reader = Excel.FromCsv(ms);
            using var e = ExcelParser.FromAttributes<PersonRow>().Parse(reader).GetEnumerator();

            Assert.Throws<NotSupportedException>(e.Reset);
        }
    }
}
