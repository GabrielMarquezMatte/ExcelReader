using System.Collections;
using System.Collections.Generic;
using System.Text;
using ExcelReader.Core.Parser;
using ExcelReader.Core.Reader;
using ExcelReader.Core.ValueObjects;

namespace ExcelReader.Tests
{
    public class RefParserTests
    {
        private sealed class SaleClass
        {
            public string? Name { get; set; }
            public int Id { get; set; }
            public double Value { get; set; }
            public DateTime Date { get; set; }
        }

        private readonly ref struct SaleNamedRef
        {
            public string? Name { get; init; }
            [ExcelRequired]
            public int Id { get; init; }
            public double Value { get; init; }
            public DateTime Date { get; init; }
        }

        private readonly ref struct SaleSpanRef
        {
            public ReadOnlySpan<byte> Name { get; init; }
            public int Id { get; init; }
        }

        private sealed class UpperCaseConverter : IExcelCellConverter<string>
        {
            public bool TryConvert(in Cell cell, bool isDate1904, IFormatProvider provider, out string value)
            {
                value = cell.GetString().ToUpperInvariant();
                return true;
            }
        }

        private readonly ref struct AttributeRef
        {
            [ExcelColumn("First Name")]
            public string? FirstName { get; init; }

            [ExcelIgnore]
            public int Ignored { get; init; }

            [ExcelConverter(typeof(UpperCaseConverter))]
            public string? Shout { get; init; }
        }

        private static readonly DateTime SampleDate = DateTime.FromOADate(45292.25);

        [Fact]
        public async Task ParseNamed_SupportsColumnAliasIgnoreAndConverterAttributes()
        {
            await using var ms = await TypedWorkbook.BuildAsync(
                ["First Name", "Ignored", "Shout"],
                ["Alice", 999, "hello"]);

            using var reader = Excel.FromXlsx(ms, leaveOpen: true);
            var enumerator = RefParser.ParseNamed<AttributeRef>(reader).GetEnumerator();
            Assert.True(enumerator.MoveNext());
            AttributeRef a = enumerator.Current;
            Assert.Equal("Alice", a.FirstName);
            Assert.Equal(0, a.Ignored);
            Assert.Equal("HELLO", a.Shout);
        }

        [Fact]
        public async Task ParseNamed_SpanProperty_BindsDirectlyToCellValue()
        {
            await using var ms = await TypedWorkbook.BuildAsync(
                ["Name", "Id"],
                ["Alice", 1]);

            using var reader = Excel.FromXlsx(ms, leaveOpen: true);
            var enumerator = RefParser.ParseNamed<SaleSpanRef>(reader).GetEnumerator();
            Assert.True(enumerator.MoveNext());
            SaleSpanRef s = enumerator.Current;
            Assert.Equal("Alice", Encoding.UTF8.GetString(s.Name));
            Assert.Equal(1, s.Id);
        }

        [Fact]
        public async Task ParseNamed_MatchesClassBasedParser()
        {
            await using var ms = await TypedWorkbook.BuildAsync(
                ["Name", "Id", "Value", "Date"],
                ["Alice", 1, 10.5, SampleDate],
                ["Bob", 2, -3.25, SampleDate.AddDays(1)]);

            using var namedReader = Excel.FromXlsx(ms, leaveOpen: true);
            var namedResults = new List<(string? Name, int Id, double Value, DateTime Date)>();
            foreach (SaleNamedRef s in RefParser.ParseNamed<SaleNamedRef>(namedReader))
            {
                namedResults.Add((s.Name, s.Id, s.Value, s.Date));
            }

            ms.Position = 0;
            using var classReader = Excel.FromXlsx(ms, leaveOpen: true);
            List<SaleClass> classResults = new ExcelParser<SaleClass>().Parse(classReader).ToList();

            Assert.Equal(2, namedResults.Count);
            Assert.Equal(classResults.Count, namedResults.Count);
            for (int i = 0; i < namedResults.Count; i++)
            {
                Assert.Equal(classResults[i].Name, namedResults[i].Name);
                Assert.Equal(classResults[i].Id, namedResults[i].Id);
                Assert.Equal(classResults[i].Value, namedResults[i].Value);
                Assert.Equal(classResults[i].Date, namedResults[i].Date);
            }
        }

        [Fact]
        public async Task ParseNamed_IsHeaderOrderIndependent()
        {
            await using var ms = await TypedWorkbook.BuildAsync(
                ["Date", "Value", "Name", "Id"],
                [SampleDate, 10.5, "Alice", 1]);

            using var reader = Excel.FromXlsx(ms, leaveOpen: true);
            var enumerator = RefParser.ParseNamed<SaleNamedRef>(reader).GetEnumerator();
            Assert.True(enumerator.MoveNext());
            SaleNamedRef s = enumerator.Current;
            Assert.Equal("Alice", s.Name);
            Assert.Equal(1, s.Id);
            Assert.Equal(10.5, s.Value);
            Assert.Equal(SampleDate, s.Date);
        }

        [Fact]
        public async Task ParseNamed_MissingRequiredColumn_Throws()
        {
            await using var ms = await TypedWorkbook.BuildAsync(
                ["Name", "Value", "Date"],
                ["Alice", 10.5, SampleDate]);

            using var reader = Excel.FromXlsx(ms, leaveOpen: true);
            var enumerable = RefParser.ParseNamed<SaleNamedRef>(reader);
            Assert.Throws<ExcelParseException>(() =>
            {
                var enumerator = enumerable.GetEnumerator();
                enumerator.MoveNext();
            });
        }

        [Fact]
        public async Task ParseNamed_EmptyCell_YieldsDefaultForThatColumn()
        {
            await using var ms = await TypedWorkbook.BuildAsync(
                ["Name", "Id", "Value", "Date"],
                ["Alice", 1, new Gap(), SampleDate]);

            using var reader = Excel.FromXlsx(ms, leaveOpen: true);
            var enumerator = RefParser.ParseNamed<SaleNamedRef>(reader).GetEnumerator();
            Assert.True(enumerator.MoveNext());
            Assert.Equal(0.0, enumerator.Current.Value);
        }

        [Fact]
        public async Task ParseNamed_RequiredColumnWithUnparseableValueThrowsAsIfMissing()
        {
            await using var ms = await TypedWorkbook.BuildAsync(
                ["Name", "Id", "Value", "Date"],
                ["Alice", "oops", 10.5, SampleDate]);

            using var reader = Excel.FromXlsx(ms, leaveOpen: true);
            var enumerator = RefParser.ParseNamed<SaleNamedRef>(reader).GetEnumerator();
            Assert.True(enumerator.MoveNext());
            Assert.Throws<ExcelParseException>(() => { _ = enumerator.Current; });
        }

        [Fact]
        public async Task ParseNamed_ThrowOnParseFailureThrowsExcelParseException()
        {
            await using var ms = await TypedWorkbook.BuildAsync(
                ["Name", "Id", "Value", "Date"],
                ["Alice", "oops", 10.5, SampleDate]);
            var config = new ExcelParserConfig { ThrowOnParseFailure = true };

            using var reader = Excel.FromXlsx(ms, leaveOpen: true);
            var enumerator = RefParser.ParseNamed<SaleNamedRef>(reader, config).GetEnumerator();
            Assert.True(enumerator.MoveNext());
            ExcelParseException ex = Assert.Throws<ExcelParseException>(() => { _ = enumerator.Current; });
            Assert.Equal("Id", ex.ColumnName);
            Assert.Equal("oops", ex.RawValue);
        }

        [Fact]
        public void ParseNamed_Date1904System_ShiftsParsedDateBy1462Days()
        {
            const string sheetRows =
                """<row r="1"><c r="A1" t="str"><v>Date</v></c><c r="B1" t="str"><v>Id</v></c></row>""" +
                """<row r="2"><c r="A2"><v>1000</v></c><c r="B2"><v>1</v></c></row>""";

            using var ms1900 = WorkbookBuilder.Build(sheetRows, date1904: false);
            using var reader1900 = Excel.FromXlsx(ms1900, leaveOpen: true);
            Assert.False(reader1900.IsDate1904);
            var e1900 = RefParser.ParseNamed<SaleNamedRef>(reader1900).GetEnumerator();
            Assert.True(e1900.MoveNext());
            DateTime date1900 = e1900.Current.Date;

            using var ms1904 = WorkbookBuilder.Build(sheetRows, date1904: true);
            using var reader1904 = Excel.FromXlsx(ms1904, leaveOpen: true);
            Assert.True(reader1904.IsDate1904);
            var e1904 = RefParser.ParseNamed<SaleNamedRef>(reader1904).GetEnumerator();
            Assert.True(e1904.MoveNext());
            DateTime date1904Result = e1904.Current.Date;

            Assert.Equal(1462, (date1904Result - date1900).Days);
        }

        [Fact]
        public async Task ParseNamed_AwaitForeach_EnumeratesRows()
        {
            await using var ms = await TypedWorkbook.BuildAsync(
                ["Name", "Id", "Value", "Date"],
                ["Alice", 1, 10.5, SampleDate],
                ["Bob", 2, -3.25, SampleDate.AddDays(1)]);

            using var reader = Excel.FromXlsx(ms, leaveOpen: true);
            var results = new List<(string? Name, int Id)>();
            await foreach (SaleNamedRef s in RefParser.ParseNamed<SaleNamedRef>(reader))
            {
                results.Add((s.Name, s.Id));
            }

            Assert.Equal(2, results.Count);
            Assert.Equal(("Alice", 1), results[0]);
            Assert.Equal(("Bob", 2), results[1]);
        }

        [Fact]
        public async Task IEnumerableInterop_Throws()
        {
            await using var ms = await TypedWorkbook.BuildAsync(["Name", "Id", "Value", "Date"]);
            using var reader = Excel.FromXlsx(ms, leaveOpen: true);
            IEnumerable<SaleNamedRef> enumerable = RefParser.ParseNamed<SaleNamedRef>(reader);
            Assert.Throws<NotSupportedException>(enumerable.GetEnumerator);
            Assert.Throws<NotSupportedException>(() => ((IEnumerable)enumerable).GetEnumerator());
        }
    }
}
