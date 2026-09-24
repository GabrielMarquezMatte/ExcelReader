using ExcelReader.Core.Parser;
using ExcelReader.Core.Reader;
using ExcelReader.Core.Reader.Xlsx;

namespace ExcelReader.Tests.Parser
{
    public partial class HeaderCollisionTests
    {
        // EXR006 fires for these two, so their Generated path is covered by GeneratorTests' in-memory compilation instead.
        private sealed class CaseModel
        {
            [ExcelColumn("Name")]
            public string Upper { get; set; } = "";
            [ExcelColumn("name")]
            public string Lower { get; set; } = "";
        }

        private sealed class AliasModel
        {
            [ExcelColumn("Id")]
            [ExcelColumn("Key")]
            public string Id { get; set; } = "";
            [ExcelColumn("key")]
            public string Other { get; set; } = "";
        }

        [ExcelSerializable]
        private sealed partial class DiacriticModel
        {
            [ExcelColumn("Café")]
            public string Accented { get; set; } = "";
            [ExcelColumn("Cafe")]
            public string Plain { get; set; } = "";
        }

        [ExcelSerializable]
        private sealed partial class SamePropertyAliasModel
        {
            [ExcelColumn("Name")]
            [ExcelColumn("NAME")]
            [ExcelColumn("Name")]
            public string Name { get; set; } = "";
        }

        private static readonly ExcelParserConfig Ordinal = new() { ColumnNameComparer = StringComparer.Ordinal };

        private static ExcelParser<CaseModel>[] CaseParsers(ExcelParserConfig? config)
        {
            return
            [
                ExcelParser.FromAttributes<CaseModel>(config),
                ExcelParser.Build<CaseModel>(static b => b
                    .Factory(static () => new CaseModel())
                    .Property(["Name"], ExcelCellReaders.String, static (ref m, v) => m.Upper = v)
                    .Property(["name"], ExcelCellReaders.String, static (ref m, v) => m.Lower = v), config),
            ];
        }

        private static ExcelParser<AliasModel>[] AliasParsers(ExcelParserConfig? config)
        {
            return
            [
                ExcelParser.FromAttributes<AliasModel>(config),
                ExcelParser.Build<AliasModel>(static b => b
                    .Factory(static () => new AliasModel())
                    .Property(["Id", "Key"], ExcelCellReaders.String, static (ref m, v) => m.Id = v)
                    .Property(["key"], ExcelCellReaders.String, static (ref m, v) => m.Other = v), config),
            ];
        }

        private static ExcelParser<DiacriticModel>[] DiacriticParsers(ExcelParserConfig? config)
        {
            return
            [
                ExcelParser.FromAttributes<DiacriticModel>(config),
                ExcelParser.Generated<DiacriticModel>(config),
                ExcelParser.Build<DiacriticModel>(static b => b
                    .Factory(static () => new DiacriticModel())
                    .Property(["Café"], ExcelCellReaders.String, static (ref m, v) => m.Accented = v)
                    .Property(["Cafe"], ExcelCellReaders.String, static (ref m, v) => m.Plain = v), config),
            ];
        }

        private static async Task<List<T>> ParseAsync<T>(ExcelParser<T> parser, object?[] header, object?[] row)
        {
            await using MemoryStream ms = await TypedWorkbook.BuildAsync(header, row);
            await using XlsxReader reader = await Excel.FromXlsxAsync(ms, ct: TestContext.Current.CancellationToken);
            return [.. parser.Parse(reader)];
        }

        [Fact]
        public async Task CaseOnlyDifferentHeadersOnTwoPropertiesThrowUnderDefaultComparer()
        {
            foreach (ExcelParser<CaseModel> parser in CaseParsers(null))
            {
                InvalidOperationException ex = await Assert.ThrowsAsync<InvalidOperationException>(
                    () => ParseAsync(parser, ["Name"], ["A"]));
                Assert.Contains("'Name'", ex.Message, StringComparison.Ordinal);
                Assert.Contains("'name'", ex.Message, StringComparison.Ordinal);
            }
        }

        [Fact]
        public async Task CaseOnlyDifferentHeadersBindSeparatelyUnderOrdinalComparer()
        {
            foreach (ExcelParser<CaseModel> parser in CaseParsers(Ordinal))
            {
                CaseModel result = Assert.Single(await ParseAsync(parser, ["Name", "name"], ["A", "b"]));
                Assert.Equal("A", result.Upper);
                Assert.Equal("b", result.Lower);
            }
        }

        [Fact]
        public async Task AliasCollidingWithAnotherPropertyThrows()
        {
            foreach (ExcelParser<AliasModel> parser in AliasParsers(null))
            {
                InvalidOperationException ex = await Assert.ThrowsAsync<InvalidOperationException>(
                    () => ParseAsync(parser, ["Id"], ["1"]));
                Assert.Contains("'key'", ex.Message, StringComparison.Ordinal);
            }
            foreach (ExcelParser<AliasModel> parser in AliasParsers(Ordinal))
            {
                AliasModel result = Assert.Single(await ParseAsync(parser, ["Key", "key"], ["1", "2"]));
                Assert.Equal("1", result.Id);
                Assert.Equal("2", result.Other);
            }
        }

        [Fact]
        public async Task NormalizationEquivalentHeadersThrowOnlyWhenNormalizationMakesThemEqual()
        {
            foreach (ExcelParser<DiacriticModel> parser in DiacriticParsers(new ExcelParserConfig { HeaderNormalization = HeaderNormalization.RemoveDiacritics }))
            {
                await Assert.ThrowsAsync<InvalidOperationException>(() => ParseAsync(parser, ["Cafe"], ["x"]));
            }
            foreach (ExcelParser<DiacriticModel> parser in DiacriticParsers(null))
            {
                DiacriticModel result = Assert.Single(await ParseAsync(parser, ["Café", "Cafe"], ["a", "b"]));
                Assert.Equal("a", result.Accented);
                Assert.Equal("b", result.Plain);
            }
        }

        [Fact]
        public async Task DuplicateAliasesOnOnePropertyAreNotAmbiguous()
        {
            ExcelParser<SamePropertyAliasModel>[] parsers =
            [
                ExcelParser.FromAttributes<SamePropertyAliasModel>(),
                ExcelParser.Generated<SamePropertyAliasModel>(),
                ExcelParser.Build<SamePropertyAliasModel>(static b => b
                    .Factory(static () => new SamePropertyAliasModel())
                    .Property(["Name", "NAME", "Name"], ExcelCellReaders.String, static (ref m, v) => m.Name = v)),
            ];
            foreach (ExcelParser<SamePropertyAliasModel> parser in parsers)
            {
                SamePropertyAliasModel result = Assert.Single(await ParseAsync(parser, ["NAME"], ["x"]));
                Assert.Equal("x", result.Name);
            }
        }

        [Fact]
        public async Task FluentOverrideOfAttributeHeaderIsNotACollision()
        {
            ExcelParser<CaseModel> parser = ExcelParser.BuildWithAttributeFallback<CaseModel>(static b => b
                .Property(["Name"], ExcelCellReaders.String, static (ref m, v) => m.Upper = v.ToUpperInvariant()), Ordinal);

            CaseModel result = Assert.Single(await ParseAsync(parser, ["Name", "name"], ["a", "b"]));
            Assert.Equal("A", result.Upper);
            Assert.Equal("b", result.Lower);
        }
    }
}
