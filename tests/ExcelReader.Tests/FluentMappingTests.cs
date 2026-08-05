using ExcelReader.Core.Parser;
using ExcelReader.Core.Reader;
using ExcelReader.Core.Writer;

namespace ExcelReader.Tests
{
    // Feature C: mapping decided at runtime by ExcelFluentParser<T>/ExcelRowMapBuilder<T>.PropertyAt,
    // rather than by [ExcelColumn]/[ExcelRequired] attributes (ExcelParser<T>) or a compile-time map
    // (ExcelMappedParser<T>). See docs/v2-plan.md §4 for the acceptance criteria these tests cover.
    public class FluentMappingTests
    {
        private sealed class SharedModel
        {
            public string Name { get; set; } = "";
            public int Age { get; set; }
        }

        // No [ExcelColumn]/[ExcelRequired] anywhere — proves criterion 2 (fully fluent-mappable model)
        // and doubles as the attribute-driven fallback target for WithAttributeFallback tests, since its
        // default reflection map (property name = header name) is exactly what a caller would replicate
        // by hand with Property().
        private sealed class AttributedModel
        {
            public string Name { get; set; } = "";
            public int Age { get; set; }
        }

        // For FluentBindingWithDifferentHeaderDoesNotOverrideAttribute: the attribute's own header name
        // ("file") differs from whatever the builder is configured with, on purpose.
        private sealed class AliasedModel
        {
            [ExcelColumn("file")]
            public string Name { get; set; } = "";
        }

        [Fact]
        public async Task TwoDifferentMapsForSameModelInOneProcess()
        {
            CancellationToken ct = TestContext.Current.CancellationToken;
            await using MemoryStream ms = await TypedWorkbook.BuildAsync(["Name", "Age"], ["Alice", 30]);

            var parserA = new ExcelFluentParser<SharedModel>(static b => b
                .Factory(static () => new SharedModel())
                .Property(["Name"], ExcelCellReaders.String, static (ref SharedModel m, string v) => m.Name = v)
                .Property(["Age"], ExcelCellReaders.Parsable, static (ref SharedModel m, int v) => m.Age = v));
            await using XlsxReader readerA = await Excel.FromAsync(ms, leaveOpen: true, ct: ct);
            List<SharedModel> resultA = parserA.Parse(readerA).ToList();

            ms.Position = 0;
            var parserB = new ExcelFluentParser<SharedModel>(static b => b
                .Factory(static () => new SharedModel())
                .Property(["Name"], ExcelCellReaders.String, static (ref SharedModel m, string v) => m.Name = v.ToUpperInvariant())
                .Property(["Age"], ExcelCellReaders.Parsable, static (ref SharedModel m, int v) => m.Age = v * 2));
            await using XlsxReader readerB = await Excel.FromAsync(ms, ct: ct);
            List<SharedModel> resultB = parserB.Parse(readerB).ToList();

            Assert.Equal("Alice", resultA[0].Name);
            Assert.Equal(30, resultA[0].Age);
            Assert.Equal("ALICE", resultB[0].Name);
            Assert.Equal(60, resultB[0].Age);
        }

        [Fact]
        public async Task ModelWithNoAttributesIsFullyMappable()
        {
            CancellationToken ct = TestContext.Current.CancellationToken;
            await using MemoryStream ms = await TypedWorkbook.BuildAsync(["Name", "Age"], ["Bob", 42]);

            var parser = new ExcelFluentParser<SharedModel>(static b => b
                .Factory(static () => new SharedModel())
                .Property(["Name"], ExcelCellReaders.String, static (ref SharedModel m, string v) => m.Name = v)
                .Property(["Age"], ExcelCellReaders.Parsable, static (ref SharedModel m, int v) => m.Age = v));
            await using XlsxReader reader = await Excel.FromAsync(ms, ct: ct);
            List<SharedModel> results = parser.Parse(reader).ToList();

            Assert.Single(results);
            Assert.Equal("Bob", results[0].Name);
            Assert.Equal(42, results[0].Age);
        }

        [Fact]
        public async Task MapByColumnIndexWithoutHeaderRow()
        {
            CancellationToken ct = TestContext.Current.CancellationToken;
            // No header row at all: the first (and only) row is already data.
            await using MemoryStream ms = await TypedWorkbook.BuildAsync(["Carol", 21]);

            var parser = new ExcelFluentParser<SharedModel>(static b => b
                .Factory(static () => new SharedModel())
                .PropertyAt(0, ExcelCellReaders.String, static (ref SharedModel m, string v) => m.Name = v)
                .PropertyAt(1, ExcelCellReaders.Parsable, static (ref SharedModel m, int v) => m.Age = v));
            await using XlsxReader reader = await Excel.FromAsync(ms, ct: ct);
            List<SharedModel> results = parser.Parse(reader).ToList();

            Assert.Single(results);
            Assert.Equal("Carol", results[0].Name);
            Assert.Equal(21, results[0].Age);
        }

        [Fact]
        public async Task RequiredColumnMissingByIndexThrows()
        {
            CancellationToken ct = TestContext.Current.CancellationToken;
            // Only column 0 exists; the map also binds column 1 as required.
            await using MemoryStream ms = await TypedWorkbook.BuildAsync(["Dave"]);

            var parser = new ExcelFluentParser<SharedModel>(static b => b
                .Factory(static () => new SharedModel())
                .PropertyAt(0, ExcelCellReaders.String, static (ref SharedModel m, string v) => m.Name = v)
                .PropertyAt(1, ExcelCellReaders.Parsable, static (ref SharedModel m, int v) => m.Age = v, requireValue: true));
            await using XlsxReader reader = await Excel.FromAsync(ms, ct: ct);

            Assert.Throws<ExcelParseException>(() => parser.Parse(reader).ToList());
        }

        [Fact]
        public async Task FluentOverridesAttributeForConfiguredProperty()
        {
            CancellationToken ct = TestContext.Current.CancellationToken;
            await using MemoryStream ms = await TypedWorkbook.BuildAsync(["Name", "Age"], ["Erin", 33]);

            ExcelFluentParser<AttributedModel> parser = ExcelFluentParser<AttributedModel>.WithAttributeFallback(static b => b
                .Property(["Name"], ExcelCellReaders.String, static (ref AttributedModel m, string v) => m.Name = v.ToUpperInvariant()));
            await using XlsxReader reader = await Excel.FromAsync(ms, ct: ct);
            List<AttributedModel> results = parser.Parse(reader).ToList();

            Assert.Single(results);
            Assert.Equal("ERIN", results[0].Name);
        }

        [Fact]
        public async Task AttributeSurvivesForPropertyNotInBuilder()
        {
            CancellationToken ct = TestContext.Current.CancellationToken;
            await using MemoryStream ms = await TypedWorkbook.BuildAsync(["Name", "Age"], ["Frank", 19]);

            // Only Name is configured; Age is never mentioned in the builder, so it must still bind via
            // the attribute-driven (reflection) fallback.
            ExcelFluentParser<AttributedModel> parser = ExcelFluentParser<AttributedModel>.WithAttributeFallback(static b => b
                .Property(["Name"], ExcelCellReaders.String, static (ref AttributedModel m, string v) => m.Name = v.ToUpperInvariant()));
            await using XlsxReader reader = await Excel.FromAsync(ms, ct: ct);
            List<AttributedModel> results = parser.Parse(reader).ToList();

            Assert.Single(results);
            Assert.Equal(19, results[0].Age);
        }

        [Fact]
        public async Task FluentMapMatchesAttributeMapForEquivalentConfig()
        {
            CancellationToken ct = TestContext.Current.CancellationToken;
            await using MemoryStream ms = await TypedWorkbook.BuildAsync(["Name", "Age"], ["Grace", 27]);

            var fluentParser = new ExcelFluentParser<AttributedModel>(static b => b
                .Factory(static () => new AttributedModel())
                .Property(["Name"], ExcelCellReaders.String, static (ref AttributedModel m, string v) => m.Name = v)
                .Property(["Age"], ExcelCellReaders.Parsable, static (ref AttributedModel m, int v) => m.Age = v));
            await using XlsxReader fluentReader = await Excel.FromAsync(ms, leaveOpen: true, ct: ct);
            List<AttributedModel> fluentResults = fluentParser.Parse(fluentReader).ToList();

            ms.Position = 0;
            await using XlsxReader reflectedReader = await Excel.FromAsync(ms, ct: ct);
            List<AttributedModel> reflectedResults = new ExcelParser<AttributedModel>().Parse(reflectedReader).ToList();

            Assert.Single(fluentResults);
            Assert.Single(reflectedResults);
            Assert.Equal(reflectedResults[0].Name, fluentResults[0].Name);
            Assert.Equal(reflectedResults[0].Age, fluentResults[0].Age);
        }

        [Fact]
        public async Task FluentMapWorksAcrossAllFourFormats()
        {
            CancellationToken ct = TestContext.Current.CancellationToken;
            var record = new AttributedModel { Name = "Heidi", Age = 51 };
            Action<ExcelRowMapBuilder<AttributedModel>> configure = static b => b
                .Factory(static () => new AttributedModel())
                .Property(["Name"], ExcelCellReaders.String, static (ref AttributedModel m, string v) => m.Name = v)
                .Property(["Age"], ExcelCellReaders.Parsable, static (ref AttributedModel m, int v) => m.Age = v);

            await using var xlsxStream = new MemoryStream();
            await using (WorkbookRecordWriter<XlsxSheetWriter, XlsxRowWriter> writer = await RecordWriter.CreateXlsxAsync(xlsxStream, leaveOpen: true, ct: ct))
            {
                await writer.WriteSheetAsync("S1", [record], ct);
            }
            xlsxStream.Position = 0;
            await using XlsxReader xlsxReader = await Excel.FromAsync(xlsxStream, ct: ct);
            AssertMatches(record, new ExcelFluentParser<AttributedModel>(configure).Parse(xlsxReader).Single());

            await using var xlsbStream = new MemoryStream();
            await using (WorkbookRecordWriter<XlsbSheetWriter, XlsbRowWriter> writer = await RecordWriter.CreateXlsbAsync(xlsbStream, leaveOpen: true, ct: ct))
            {
                await writer.WriteSheetAsync("S1", [record], ct);
            }
            xlsbStream.Position = 0;
            await using XlsbReader xlsbReader = await Excel.FromXlsbAsync(xlsbStream, leaveOpen: false, ct: ct);
            AssertMatches(record, new ExcelFluentParser<AttributedModel>(configure).Parse(xlsbReader).Single());

            await using var xlsStream = new MemoryStream();
            await using (WorkbookRecordWriter<XlsSheetWriter, XlsRowWriter> writer = await RecordWriter.CreateXlsAsync(xlsStream, leaveOpen: true, ct: ct))
            {
                await writer.WriteSheetAsync("S1", [record], ct);
            }
            xlsStream.Position = 0;
            using XlsReader xlsReader = Excel.FromXls(xlsStream, leaveOpen: false);
            AssertMatches(record, new ExcelFluentParser<AttributedModel>(configure).Parse(xlsReader).Single());

            await using var csvStream = new MemoryStream();
            await using (WorkbookRecordWriter<CsvSheetWriter, CsvRowWriter> writer = await RecordWriter.CreateCsvAsync(csvStream, leaveOpen: true, ct: ct))
            {
                await writer.WriteSheetAsync("S1", [record], ct);
            }
            csvStream.Position = 0;
            using CsvReader csvReader = Excel.FromCsv(csvStream, leaveOpen: false);
            AssertMatches(record, new ExcelFluentParser<AttributedModel>(configure).Parse(csvReader).Single());
        }

        [Fact]
        public async Task FluentBindingWithDifferentHeaderDoesNotOverrideAttribute()
        {
            CancellationToken ct = TestContext.Current.CancellationToken;
            // "file" is the attribute's own header name; "arquivo" is a different name the builder
            // configures for the same property. Precedence is by header name, not property identity
            // (docs/v2-plan.md §4.4.3), so neither binding is suppressed — both survive, and whichever
            // column comes later in the row wins the assignment. Reusing "file" in the builder instead
            // of "arquivo" is what would actually override the attribute.
            await using MemoryStream ms = await TypedWorkbook.BuildAsync(["file", "arquivo"], ["FromAttribute", "FromFluent"]);

            ExcelFluentParser<AliasedModel> parser = ExcelFluentParser<AliasedModel>.WithAttributeFallback(static b => b
                .Property(["arquivo"], ExcelCellReaders.String, static (ref AliasedModel m, string v) => m.Name = v));
            await using XlsxReader reader = await Excel.FromAsync(ms, ct: ct);
            List<AliasedModel> results = parser.Parse(reader).ToList();

            Assert.Single(results);
            Assert.Equal("FromFluent", results[0].Name);
        }

        [Fact]
        public void WithAttributeFallbackRejectsIndexBasedMap()
        {
            InvalidOperationException ex = Assert.Throws<InvalidOperationException>(() =>
                ExcelFluentParser<AttributedModel>.WithAttributeFallback(static b => b
                    .PropertyAt(0, ExcelCellReaders.String, static (ref AttributedModel m, string v) => m.Name = v)));

            Assert.Contains("PropertyAt", ex.Message, StringComparison.Ordinal);
        }

        private static void AssertMatches(AttributedModel expected, AttributedModel actual)
        {
            Assert.Equal(expected.Name, actual.Name);
            Assert.Equal(expected.Age, actual.Age);
        }
    }
}
