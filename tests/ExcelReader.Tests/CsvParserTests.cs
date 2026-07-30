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
        public void HigherPriorityAliasReplacesEarlierLowerPriorityBinding()
        {
            // "Legacy Name" (alias index 1) binds first since it's the earlier column; "Preferred
            // Name" (alias index 0) then takes over the property, unbinding the earlier column.
            using var ms = Csv("Legacy Name,Preferred Name\nOld,New\n");
            using var reader = Excel.FromCsv(ms);

            MultiAliasRow row = new ExcelParser<MultiAliasRow>().Parse(reader).Single();

            Assert.Equal("New", row.Name);
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
        public void EnumFromNumericTextBindsByValueInCsv()
        {
            // Every CSV cell is text, so an enum written as its underlying number ("2") arrives as
            // text. The enum name map registers each member's numeric string form alongside its name,
            // so numeric-text enum columns resolve in CSV too.
            var id = Guid.NewGuid();
            using var ms = Csv($"Status,Id,Quantity\n2,{id},7\n");
            using var reader = Excel.FromCsv(ms);

            TypedRow row = new ExcelParser<TypedRow>().Parse(reader).Single();

            Assert.Equal(Status.Closed, row.Status);
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

            Assert.Throws<ExcelParseException>(() => new ExcelParser<RequiredRow>().Parse(reader).ToList());
        }

        [Fact]
        public void EmptyRequiredCellThrowsNamingColumnAndRow()
        {
            using var ms = Csv("Id,Note\n,hi\n");
            using var reader = Excel.FromCsv(ms);

            Assert.Throws<ExcelParseException>(() => new ExcelParser<RequiredRow>().Parse(reader).ToList());
        }

        [Fact]
        public void RequiredCellWithUnparseableValueThrowsAsIfMissing()
        {
            // "Id" is present and non-empty but "abc" isn't a valid int — F3: treated the same as a
            // blank required cell instead of silently leaving the model's Id at 0.
            using var ms = Csv("Id,Note\nabc,hi\n");
            using var reader = Excel.FromCsv(ms);

            ExcelParseException ex = Assert.Throws<ExcelParseException>(
                () => new ExcelParser<RequiredRow>().Parse(reader).ToList());
            Assert.Contains("Id", ex.Message, StringComparison.Ordinal);
        }

        [Fact]
        public void ThrowOnParseFailureThrowsExcelParseExceptionForUnparseableColumn()
        {
            using var ms = Csv("Name,Amount\nConta,not-a-number\n");
            using var reader = Excel.FromCsv(ms);
            var config = new ExcelParserConfig { ThrowOnParseFailure = true };

            ExcelParseException ex = Assert.Throws<ExcelParseException>(
                () => new ExcelParser<MoneyRow>(config).Parse(reader).ToList());
            Assert.Equal("Amount", ex.ColumnName);
            Assert.Equal("not-a-number", ex.RawValue);
        }

        [Fact]
        public void TerminalBlankLineDoesNotYieldPhantomModelOrRequiredFailure()
        {
            using var ms = Csv("Id,Note\n7,valid\n\n");
            using var reader = Excel.FromCsv(ms);

            RequiredRow row = Assert.Single(new ExcelParser<RequiredRow>().Parse(reader).ToList());

            Assert.Equal(7, row.Id);
            Assert.Equal("valid", row.Note);
        }

        [Fact]
        public void PlainDateTimeAndDateOnlyColumnsParseTextNatively()
        {
            // The CSV parser parses DateTime/DateOnly from the cell text (no [ExcelConverter] needed) —
            // unlike the Excel readers, where those types interpret an Excel serial number.
            using var ms = Csv("Created,Day\n2026-07-02T08:30:00,2026-07-02\n");
            using var reader = Excel.FromCsv(ms);

            DateRow row = new ExcelParser<DateRow>().Parse(reader).Single();

            Assert.Equal(new DateTime(2026, 7, 2, 8, 30, 0, DateTimeKind.Unspecified), row.Created);
            Assert.Equal(new DateOnly(2026, 7, 2), row.Day);
        }

        [Fact]
        public void UnparseableDateColumnKeepsDefault()
        {
            using var ms = Csv("Created,Day\nnot-a-date,\n");
            using var reader = Excel.FromCsv(ms);

            DateRow row = new ExcelParser<DateRow>().Parse(reader).Single();

            Assert.Equal(default, row.Created);
            Assert.Null(row.Day);
        }

        [Fact]
        public void DateColumnRespectsCulture()
        {
            // pt-BR day-first format: "02/07/2026" is 2 July, not 7 February.
            using var ms = Csv("Created,Day\n02/07/2026,02/07/2026\n");
            using var reader = Excel.FromCsv(ms);
            var config = new ExcelParserConfig { Culture = CultureInfo.GetCultureInfo("pt-BR") };

            DateRow row = new ExcelParser<DateRow>(config).Parse(reader).Single();

            Assert.Equal(new DateTime(2026, 7, 2, 0, 0, 0, DateTimeKind.Unspecified), row.Created);
            Assert.Equal(new DateOnly(2026, 7, 2), row.Day);
        }

        [Fact]
        public void CustomConverterStillOverridesDateParsing()
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

        [Fact]
        public void NonGenericEnumerableGetEnumeratorWorks()
        {
            using var ms = Csv("Name\nAlice\n");
            using var reader = Excel.FromCsv(ms);
            IEnumerable enumerable = new ExcelParser<PersonRow>().Parse(reader);

            IEnumerator e = enumerable.GetEnumerator();

            Assert.True(e.MoveNext());
            Assert.Equal("Alice", Assert.IsType<PersonRow>(e.Current).Name);
        }

        [Fact]
        public void EnumeratorResetThrows()
        {
            using var ms = Csv("Name\nAlice\n");
            using var reader = Excel.FromCsv(ms);
            using var e = new ExcelParser<PersonRow>().Parse(reader).GetEnumerator();

            Assert.Throws<NotSupportedException>(e.Reset);
        }
    }
}
