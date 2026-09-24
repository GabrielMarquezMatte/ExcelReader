using ExcelReader.Core.Parser;
using ExcelReader.Core.Reader;
using ExcelReader.Core.Reader.Csv;
using ExcelReader.Core.Reader.Xls;
using ExcelReader.Core.Reader.Xlsb;
using ExcelReader.Core.Reader.Xlsx;
using ExcelReader.Core.Writer;
using ExcelReader.Core.Writer.Csv;
using ExcelReader.Core.Writer.Xls;
using ExcelReader.Core.Writer.Xlsb;
using ExcelReader.Core.Writer.Xlsx;

namespace ExcelReader.Tests.Parser
{
    public class FluentMappingTests
    {
        private sealed class SharedModel
        {
            public string Name { get; set; } = "";
            public int Age { get; set; }
        }

        private sealed class AttributedModel
        {
            public string Name { get; set; } = "";
            public int Age { get; set; }
        }

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

            var parserA = ExcelParser.Build<SharedModel>(static b => b
                .Factory(static () => new SharedModel())
                .Property(["Name"], ExcelCellReaders.String, static (ref m, v) => m.Name = v)
                .Property(["Age"], ExcelCellReaders.Parsable, static (ref SharedModel m, int v) => m.Age = v));
            await using XlsxReader readerA = await Excel.FromXlsxAsync(ms, leaveOpen: true, ct: ct);
            List<SharedModel> resultA = parserA.Parse(readerA).ToList();

            ms.Position = 0;
            var parserB = ExcelParser.Build<SharedModel>(static b => b
                .Factory(static () => new SharedModel())
                .Property(["Name"], ExcelCellReaders.String, static (ref m, v) => m.Name = v.ToUpperInvariant())
                .Property(["Age"], ExcelCellReaders.Parsable, static (ref SharedModel m, int v) => m.Age = v * 2));
            await using XlsxReader readerB = await Excel.FromXlsxAsync(ms, ct: ct);
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

            var parser = ExcelParser.Build<SharedModel>(static b => b
                .Factory(static () => new SharedModel())
                .Property(["Name"], ExcelCellReaders.String, static (ref m, v) => m.Name = v)
                .Property(["Age"], ExcelCellReaders.Parsable, static (ref SharedModel m, int v) => m.Age = v));
            await using XlsxReader reader = await Excel.FromXlsxAsync(ms, ct: ct);
            List<SharedModel> results = parser.Parse(reader).ToList();

            Assert.Single(results);
            Assert.Equal("Bob", results[0].Name);
            Assert.Equal(42, results[0].Age);
        }

        [Fact]
        public async Task MapByColumnIndexWithoutHeaderRow()
        {
            CancellationToken ct = TestContext.Current.CancellationToken;
            await using MemoryStream ms = await TypedWorkbook.BuildAsync(["Carol", 21]);

            var parser = ExcelParser.Build<SharedModel>(static b => b
                .Factory(static () => new SharedModel())
                .PropertyAt(0, ExcelCellReaders.String, static (ref m, v) => m.Name = v)
                .PropertyAt(1, ExcelCellReaders.Parsable, static (ref SharedModel m, int v) => m.Age = v));
            await using XlsxReader reader = await Excel.FromXlsxAsync(ms, ct: ct);
            List<SharedModel> results = parser.Parse(reader).ToList();

            Assert.Single(results);
            Assert.Equal("Carol", results[0].Name);
            Assert.Equal(21, results[0].Age);
        }

        [Fact]
        public async Task RequiredColumnMissingByIndexThrows()
        {
            CancellationToken ct = TestContext.Current.CancellationToken;
            await using MemoryStream ms = await TypedWorkbook.BuildAsync(["Dave"]);

            var parser = ExcelParser.Build<SharedModel>(static b => b
                .Factory(static () => new SharedModel())
                .PropertyAt(0, ExcelCellReaders.String, static (ref m, v) => m.Name = v)
                .PropertyAt(1, ExcelCellReaders.Parsable, static (ref SharedModel m, int v) => m.Age = v, requireValue: true));
            await using XlsxReader reader = await Excel.FromXlsxAsync(ms, ct: ct);

            Assert.Throws<ExcelParseException>(() => parser.Parse(reader).ToList());
        }

        [Fact]
        public async Task FluentOverridesAttributeForConfiguredProperty()
        {
            CancellationToken ct = TestContext.Current.CancellationToken;
            await using MemoryStream ms = await TypedWorkbook.BuildAsync(["Name", "Age"], ["Erin", 33]);

            ExcelParser<AttributedModel> parser = ExcelParser.BuildWithAttributeFallback<AttributedModel>(static b => b
                .Property(["Name"], ExcelCellReaders.String, static (ref m, v) => m.Name = v.ToUpperInvariant()));
            await using XlsxReader reader = await Excel.FromXlsxAsync(ms, ct: ct);
            List<AttributedModel> results = parser.Parse(reader).ToList();

            Assert.Single(results);
            Assert.Equal("ERIN", results[0].Name);
        }

        [Fact]
        public async Task AttributeSurvivesForPropertyNotInBuilder()
        {
            CancellationToken ct = TestContext.Current.CancellationToken;
            await using MemoryStream ms = await TypedWorkbook.BuildAsync(["Name", "Age"], ["Frank", 19]);

            ExcelParser<AttributedModel> parser = ExcelParser.BuildWithAttributeFallback<AttributedModel>(static b => b
                .Property(["Name"], ExcelCellReaders.String, static (ref m, v) => m.Name = v.ToUpperInvariant()));
            await using XlsxReader reader = await Excel.FromXlsxAsync(ms, ct: ct);
            List<AttributedModel> results = parser.Parse(reader).ToList();

            Assert.Single(results);
            Assert.Equal(19, results[0].Age);
        }

        [Fact]
        public async Task FluentMapMatchesAttributeMapForEquivalentConfig()
        {
            CancellationToken ct = TestContext.Current.CancellationToken;
            await using MemoryStream ms = await TypedWorkbook.BuildAsync(["Name", "Age"], ["Grace", 27]);

            var fluentParser = ExcelParser.Build<AttributedModel>(static b => b
                .Factory(static () => new AttributedModel())
                .Property(["Name"], ExcelCellReaders.String, static (ref m, v) => m.Name = v)
                .Property(["Age"], ExcelCellReaders.Parsable, static (ref AttributedModel m, int v) => m.Age = v));
            await using XlsxReader fluentReader = await Excel.FromXlsxAsync(ms, leaveOpen: true, ct: ct);
            List<AttributedModel> fluentResults = fluentParser.Parse(fluentReader).ToList();

            ms.Position = 0;
            await using XlsxReader reflectedReader = await Excel.FromXlsxAsync(ms, ct: ct);
            List<AttributedModel> reflectedResults = ExcelParser.FromAttributes<AttributedModel>().Parse(reflectedReader).ToList();

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
                .Property(["Name"], ExcelCellReaders.String, static (ref m, v) => m.Name = v)
                .Property(["Age"], ExcelCellReaders.Parsable, static (ref AttributedModel m, int v) => m.Age = v);

            await using var xlsxStream = new MemoryStream();
            await using (XlsxWorkbookWriter writer = XlsxWorkbookWriter.Create(xlsxStream, leaveOpen: true))
            {
                await using (var sheet = writer.AddSheet("S1"))
                {
                    await sheet.WriteRecordsAsync([record], ExcelRecordLayout.FromAttributes<AttributedModel>(), ct);
                }
            }
            xlsxStream.Position = 0;
            await using XlsxReader xlsxReader = await Excel.FromXlsxAsync(xlsxStream, ct: ct);
            AssertMatches(record, ExcelParser.Build(configure).Parse(xlsxReader).Single());

            await using var xlsbStream = new MemoryStream();
            await using (XlsbWorkbookWriter writer = XlsbWorkbookWriter.Create(xlsbStream, leaveOpen: true))
            {
                await using (var sheet = writer.AddSheet("S1"))
                {
                    await sheet.WriteRecordsAsync([record], ExcelRecordLayout.FromAttributes<AttributedModel>(), ct);
                }
            }
            xlsbStream.Position = 0;
            await using XlsbReader xlsbReader = await Excel.FromXlsbAsync(xlsbStream, leaveOpen: false, ct: ct);
            AssertMatches(record, ExcelParser.Build(configure).Parse(xlsbReader).Single());

            await using var xlsStream = new MemoryStream();
            await using (XlsWorkbookWriter writer = XlsWorkbookWriter.Create(xlsStream, leaveOpen: true))
            {
                await using (var sheet = writer.AddSheet("S1"))
                {
                    await sheet.WriteRecordsAsync([record], ExcelRecordLayout.FromAttributes<AttributedModel>(), ct);
                }
            }
            xlsStream.Position = 0;
            using XlsReader xlsReader = Excel.FromXls(xlsStream, leaveOpen: false);
            AssertMatches(record, ExcelParser.Build(configure).Parse(xlsReader).Single());

            await using var csvStream = new MemoryStream();
            await using (CsvWorkbookWriter writer = CsvWorkbookWriter.Create(csvStream, leaveOpen: true))
            {
                await using (var sheet = writer.AddSheet("S1"))
                {
                    await sheet.WriteRecordsAsync([record], ExcelRecordLayout.FromAttributes<AttributedModel>(), ct);
                }
            }
            csvStream.Position = 0;
            using CsvReader csvReader = Excel.FromCsv(csvStream, leaveOpen: false);
            AssertMatches(record, ExcelParser.Build(configure).Parse(csvReader).Single());
        }

        [Fact]
        public async Task FluentBindingWithDifferentHeaderDoesNotOverrideAttribute()
        {
            CancellationToken ct = TestContext.Current.CancellationToken;
            await using MemoryStream ms = await TypedWorkbook.BuildAsync(["file", "arquivo"], ["FromAttribute", "FromFluent"]);

            ExcelParser<AliasedModel> parser = ExcelParser.BuildWithAttributeFallback<AliasedModel>(static b => b
                .Property(["arquivo"], ExcelCellReaders.String, static (ref m, v) => m.Name = v));
            await using XlsxReader reader = await Excel.FromXlsxAsync(ms, ct: ct);
            List<AliasedModel> results = parser.Parse(reader).ToList();

            Assert.Single(results);
            Assert.Equal("FromFluent", results[0].Name);
        }

        [Fact]
        public void WithAttributeFallbackRejectsIndexBasedMap()
        {
            InvalidOperationException ex = Assert.Throws<InvalidOperationException>(() =>
                ExcelParser.BuildWithAttributeFallback<AttributedModel>(static b => b
                    .PropertyAt(0, ExcelCellReaders.String, static (ref m, v) => m.Name = v)));

            Assert.Contains("PropertyAt", ex.Message, StringComparison.Ordinal);
        }

        [Fact]
        public void PlainConstructorRejectsReferenceTypeModelWithNoFactory()
        {
            InvalidOperationException ex = Assert.Throws<InvalidOperationException>(() =>
                ExcelParser.Build<AttributedModel>(static b => b
                    .Property(["Name"], ExcelCellReaders.String, static (ref m, v) => m.Name = v)));

            Assert.Contains("Factory", ex.Message, StringComparison.Ordinal);
        }

        [Fact]
        public async Task PlainConstructorAllowsValueTypeModelWithNoFactory()
        {
            CancellationToken ct = TestContext.Current.CancellationToken;
            await using MemoryStream ms = await TypedWorkbook.BuildAsync(["Age"], [42]);

            var parser = ExcelParser.Build<StructModel>(static b => b
                .Property(["Age"], ExcelCellReaders.Parsable, static (ref StructModel m, int v) => m.Age = v));
            await using XlsxReader reader = await Excel.FromXlsxAsync(ms, ct: ct);
            List<StructModel> results = parser.Parse(reader).ToList();

            Assert.Single(results);
            Assert.Equal(42, results[0].Age);
        }

        private struct StructModel
        {
            public int Age { get; set; }
        }

        private static void AssertMatches(AttributedModel expected, AttributedModel actual)
        {
            Assert.Equal(expected.Name, actual.Name);
            Assert.Equal(expected.Age, actual.Age);
        }
    }
}
