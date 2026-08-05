using ExcelReader.Core.Parser;
using ExcelReader.Core.Reader;
using ExcelReader.Core.Writer;
using ExcelReader.Core.ValueObjects;

namespace ExcelReader.Tests
{
    // Read-side parity between the reflection path (ExcelParser<T>) and the source-generated path
    // (ExcelMappedParser<T>, requiring [ExcelSerializable]) — the generator's own tabela de decisão
    // (type -> reader) is a second, independent copy of ColumnParserFactory's, and the two can diverge
    // silently (see GeneratorTests.InheritedPropertiesAreMapped's history). Every model here is declared
    // directly in this file, not as a string literal fed to CSharpGeneratorDriver, so the real
    // build-time analyzer (wired into ExcelReader.Tests.csproj with OutputItemType="Analyzer") produces
    // its IExcelRowMap<T>/IExcelRecordMap<T> for it — ExcelMappedParser<T> resolves as ordinary
    // compile-time generics, no reflection-over-a-synthetic-assembly gymnastics needed.
    //
    // Every fixture here is built with TypedWorkbook/raw writers rather than MappedRecordWriter, so this
    // file exercises only the read side; write-side parity is GeneratedRecordMapMatchesReflectionOnWrite
    // in WriterStyleTests-adjacent coverage (a separate task).
    public partial class GeneratedMapParityTests
    {
        public enum ParityKind { Alpha, Beta, Gamma }

        [ExcelSerializable]
        public partial class EveryTypeModel
        {
            public string Name { get; set; } = "";
            public bool Active { get; set; }
            public DateTime BirthDate { get; set; }
            public DateOnly BirthDay { get; set; }
            public TimeOnly BirthTime { get; set; }
            public Guid Id { get; set; }
            public byte U8 { get; set; }
            public sbyte I8 { get; set; }
            public short I16 { get; set; }
            public ushort U16 { get; set; }
            public int I32 { get; set; }
            public uint U32 { get; set; }
            public long I64 { get; set; }
            public ulong U64 { get; set; }
            public float F32 { get; set; }
            public double F64 { get; set; }
            public decimal Money { get; set; }
            public ParityKind Category { get; set; }
            public bool? ActiveN { get; set; }
            public int? I32N { get; set; }
            public decimal? MoneyN { get; set; }
            public ParityKind? CategoryN { get; set; }
            public Guid? IdN { get; set; }
        }

        public sealed class UpperCaseConverter : IExcelCellConverter<string>, IExcelCellWriter<string>
        {
            public bool TryConvert(in Cell cell, bool isDate1904, IFormatProvider provider, out string value)
            {
                value = cell.GetString().ToUpperInvariant();
                return true;
            }

            public void Write(IRowWriter row, string value)
            {
                row.Write(value);
            }
        }

        [ExcelSerializable]
        public partial class ConverterModel
        {
            [ExcelConverter(typeof(UpperCaseConverter))]
            public string Name { get; set; } = "";
        }

        [ExcelSerializable]
        public partial class AliasModel
        {
            [ExcelColumn("Legacy Name")]
            [ExcelColumn("Name")]
            public string Name { get; set; } = "";
        }

        [ExcelSerializable]
        public partial class IgnoreModel
        {
            public string Name { get; set; } = "";
            [ExcelIgnore]
            public string Ignored { get; set; } = "default";
        }

        public class InheritedBaseRow
        {
            public string Inherited { get; set; } = "";
        }

        [ExcelSerializable]
        public partial class InheritedModel : InheritedBaseRow
        {
            public int Own { get; set; }
        }

        [ExcelSerializable]
        public partial class RequiredModel
        {
            [ExcelRequired]
            public string Name { get; set; } = "";
        }

        // Deliberately excludes DateTime/DateOnly: CsvRowWriter writes them as ISO text, but the
        // generated map always reads DateTime/DateOnly as an Excel serial number (no csvTextDates
        // equivalent yet — see docs/v2-audit-fixes.md T8). TimeOnly is unaffected: every format writes
        // it as the same numeric day-fraction, so it round-trips through CSV today.
        [ExcelSerializable]
        public partial class CrossFormatModel
        {
            public string Name { get; set; } = "";
            public bool Active { get; set; }
            public int Age { get; set; }
            public decimal Balance { get; set; }
            public ParityKind Category { get; set; }
            public TimeOnly Clock { get; set; }
            public Guid Id { get; set; }
            public int? OptionalAge { get; set; }
        }

        [Fact]
        public async Task GeneratedMapMatchesReflectionForEveryBuiltInType()
        {
            var kind = ParityKind.Beta;
            var id = Guid.Parse("11111111-2222-3333-4444-555555555555");
            object?[] row =
            [
                // BirthDay is a DateTime here, not a DateOnly: TypedWorkbook.WriteCell has no DateOnly
                // case, and the underlying cell is just a date serial number either way — ReadDateOnly
                // truncates a DateTime-typed cell the same way regardless of which typed overload wrote it.
                "Alice", true, new DateTime(2024, 5, 6), new DateTime(2024, 5, 6),
                // BirthTime is a day-fraction (0.5 = noon), matching what TimeOnly's serial actually is.
                0.5, id.ToString(), 200, -100, -30000, 60000,
                2_000_000_000, 3_000_000_000d, 9_000_000_000d, 9_000_000_000d,
                1.5, 2.5, 12345.67m, kind.ToString(),
                true, 42, 99.99m, ParityKind.Gamma.ToString(), id.ToString(),
            ];
            object?[] header =
            [
                "Name", "Active", "BirthDate", "BirthDay", "BirthTime", "Id",
                "U8", "I8", "I16", "U16", "I32", "U32", "I64", "U64",
                "F32", "F64", "Money", "Category",
                "ActiveN", "I32N", "MoneyN", "CategoryN", "IdN",
            ];
            await using var ms = await TypedWorkbook.BuildAsync(header, row);

            EveryTypeModel reflectionResult = await ParseFirstXlsxAsync(ms, static reader => new ExcelParser<EveryTypeModel>().Parse(reader));
            EveryTypeModel generatedResult = await ParseFirstXlsxAsync(ms, static reader => new ExcelMappedParser<EveryTypeModel>().Parse(reader));

            Assert.Equal(reflectionResult.Name, generatedResult.Name);
            Assert.Equal(reflectionResult.Active, generatedResult.Active);
            Assert.Equal(reflectionResult.BirthDate, generatedResult.BirthDate);
            Assert.Equal(reflectionResult.BirthDay, generatedResult.BirthDay);
            Assert.Equal(reflectionResult.BirthTime, generatedResult.BirthTime);
            Assert.Equal(reflectionResult.Id, generatedResult.Id);
            Assert.Equal(reflectionResult.U8, generatedResult.U8);
            Assert.Equal(reflectionResult.I8, generatedResult.I8);
            Assert.Equal(reflectionResult.I16, generatedResult.I16);
            Assert.Equal(reflectionResult.U16, generatedResult.U16);
            Assert.Equal(reflectionResult.I32, generatedResult.I32);
            Assert.Equal(reflectionResult.U32, generatedResult.U32);
            Assert.Equal(reflectionResult.I64, generatedResult.I64);
            Assert.Equal(reflectionResult.U64, generatedResult.U64);
            Assert.Equal(reflectionResult.F32, generatedResult.F32);
            Assert.Equal(reflectionResult.F64, generatedResult.F64);
            Assert.Equal(reflectionResult.Money, generatedResult.Money);
            Assert.Equal(reflectionResult.Category, generatedResult.Category);
            Assert.Equal(reflectionResult.ActiveN, generatedResult.ActiveN);
            Assert.Equal(reflectionResult.I32N, generatedResult.I32N);
            Assert.Equal(reflectionResult.MoneyN, generatedResult.MoneyN);
            Assert.Equal(reflectionResult.CategoryN, generatedResult.CategoryN);
            Assert.Equal(reflectionResult.IdN, generatedResult.IdN);

            // Pin the actual values too, not just "the two paths agree" — two independently wrong
            // readers could still agree with each other.
            Assert.Equal("Alice", generatedResult.Name);
            Assert.Equal(ParityKind.Beta, generatedResult.Category);
            Assert.Equal(id, generatedResult.Id);
        }

        [Fact]
        public async Task GeneratedMapMatchesReflectionForCustomConverter()
        {
            await using var ms = await TypedWorkbook.BuildAsync(["Name"], ["alice"]);

            ConverterModel reflectionResult = await ParseFirstXlsxAsync(ms, static reader => new ExcelParser<ConverterModel>().Parse(reader));
            ConverterModel generatedResult = await ParseFirstXlsxAsync(ms, static reader => new ExcelMappedParser<ConverterModel>().Parse(reader));

            Assert.Equal("ALICE", reflectionResult.Name);
            Assert.Equal("ALICE", generatedResult.Name);
        }

        [Fact]
        public async Task GeneratedMapHonorsColumnAliases()
        {
            // Header uses the second alias, not the first — proves alias matching, not just
            // "first name happens to be the header".
            await using var ms = await TypedWorkbook.BuildAsync(["Name"], ["Alice"]);

            AliasModel reflectionResult = await ParseFirstXlsxAsync(ms, static reader => new ExcelParser<AliasModel>().Parse(reader));
            AliasModel generatedResult = await ParseFirstXlsxAsync(ms, static reader => new ExcelMappedParser<AliasModel>().Parse(reader));

            Assert.Equal("Alice", reflectionResult.Name);
            Assert.Equal("Alice", generatedResult.Name);
        }

        [Fact]
        public async Task GeneratedMapHonorsIgnore()
        {
            await using var ms = await TypedWorkbook.BuildAsync(["Name", "Ignored"], ["Alice", "SHOULD_NOT_BIND"]);

            IgnoreModel reflectionResult = await ParseFirstXlsxAsync(ms, static reader => new ExcelParser<IgnoreModel>().Parse(reader));
            IgnoreModel generatedResult = await ParseFirstXlsxAsync(ms, static reader => new ExcelMappedParser<IgnoreModel>().Parse(reader));

            Assert.Equal("Alice", reflectionResult.Name);
            Assert.Equal("default", reflectionResult.Ignored);
            Assert.Equal("Alice", generatedResult.Name);
            Assert.Equal("default", generatedResult.Ignored);
        }

        [Fact]
        public async Task GeneratedMapMatchesReflectionForInheritedProperties()
        {
            await using var ms = await TypedWorkbook.BuildAsync(["Inherited", "Own"], ["BaseValue", 42]);

            InheritedModel reflectionResult = await ParseFirstXlsxAsync(ms, static reader => new ExcelParser<InheritedModel>().Parse(reader));
            InheritedModel generatedResult = await ParseFirstXlsxAsync(ms, static reader => new ExcelMappedParser<InheritedModel>().Parse(reader));

            Assert.Equal("BaseValue", reflectionResult.Inherited);
            Assert.Equal(42, reflectionResult.Own);
            Assert.Equal("BaseValue", generatedResult.Inherited);
            Assert.Equal(42, generatedResult.Own);
        }

        [Fact]
        public async Task GeneratedMapThrowsOnMissingRequiredHeader()
        {
            await using var ms = await TypedWorkbook.BuildAsync(["Other"], ["x"]);

            ExcelParseException reflectionException = await Assert.ThrowsAsync<ExcelParseException>(
                () => ParseFirstXlsxAsync(ms, static reader => new ExcelParser<RequiredModel>().Parse(reader)));
            ExcelParseException generatedException = await Assert.ThrowsAsync<ExcelParseException>(
                () => ParseFirstXlsxAsync(ms, static reader => new ExcelMappedParser<RequiredModel>().Parse(reader)));

            Assert.Contains("Name", reflectionException.ColumnName, StringComparison.Ordinal);
            Assert.Contains("Name", generatedException.ColumnName, StringComparison.Ordinal);
        }

        [Fact]
        public async Task GeneratedMapParityAcrossAllFourFormatsXlsx()
        {
            var value = SampleCrossFormatValue();
            await using var ms = new MemoryStream();
            await using (XlsxWorkbookWriter wb = await XlsxWorkbookWriter.CreateAsync(ms, leaveOpen: true, ct: TestContext.Current.CancellationToken))
            {
                await wb.StartAsync(TestContext.Current.CancellationToken);
                XlsxSheetWriter sheet = wb.AddSheet("S1");
                await sheet.StartAsync(TestContext.Current.CancellationToken);
                await using (XlsxRowWriter header = await sheet.StartRowAsync(TestContext.Current.CancellationToken))
                {
                    WriteCrossFormatHeaders(header);
                }
                await using (XlsxRowWriter row = await sheet.StartRowAsync(TestContext.Current.CancellationToken))
                {
                    WriteCrossFormatRow(row, value);
                }
                await sheet.EndAsync(TestContext.Current.CancellationToken);
            }
            ms.Position = 0;
            await using XlsxReader reflectionReader = await Excel.FromAsync(ms, ct: TestContext.Current.CancellationToken);
            CrossFormatModel reflectionResult = new ExcelParser<CrossFormatModel>().Parse(reflectionReader).First();
            ms.Position = 0;
            await using XlsxReader generatedReader = await Excel.FromAsync(ms, ct: TestContext.Current.CancellationToken);
            CrossFormatModel generatedResult = new ExcelMappedParser<CrossFormatModel>().Parse(generatedReader).First();

            AssertCrossFormatEqual(reflectionResult, generatedResult);
        }

        [Fact]
        public async Task GeneratedMapParityAcrossAllFourFormatsXlsb()
        {
            var value = SampleCrossFormatValue();
            await using var ms = new MemoryStream();
            await using (XlsbWorkbookWriter wb = await XlsbWorkbookWriter.CreateAsync(ms, leaveOpen: true, ct: TestContext.Current.CancellationToken))
            {
                await wb.StartAsync(TestContext.Current.CancellationToken);
                XlsbSheetWriter sheet = wb.AddSheet("S1");
                await sheet.StartAsync(TestContext.Current.CancellationToken);
                await using (XlsbRowWriter header = await sheet.StartRowAsync(TestContext.Current.CancellationToken))
                {
                    WriteCrossFormatHeaders(header);
                }
                await using (XlsbRowWriter row = await sheet.StartRowAsync(TestContext.Current.CancellationToken))
                {
                    WriteCrossFormatRow(row, value);
                }
                await sheet.EndAsync(TestContext.Current.CancellationToken);
            }
            ms.Position = 0;
            using XlsbReader reflectionReader = Excel.FromXlsb(ms);
            CrossFormatModel reflectionResult = new ExcelParser<CrossFormatModel>().Parse(reflectionReader).First();
            ms.Position = 0;
            using XlsbReader generatedReader = Excel.FromXlsb(ms);
            CrossFormatModel generatedResult = new ExcelMappedParser<CrossFormatModel>().Parse(generatedReader).First();

            AssertCrossFormatEqual(reflectionResult, generatedResult);
        }

        [Fact]
        public async Task GeneratedMapParityAcrossAllFourFormatsXls()
        {
            var value = SampleCrossFormatValue();
            await using var ms = new MemoryStream();
            await using (XlsWorkbookWriter wb = XlsWorkbookWriter.Create(ms, leaveOpen: true))
            {
                await wb.StartAsync(TestContext.Current.CancellationToken);
                XlsSheetWriter sheet = wb.AddSheet("S1");
                await sheet.StartAsync(TestContext.Current.CancellationToken);
                await using (XlsRowWriter header = await sheet.StartRowAsync(TestContext.Current.CancellationToken))
                {
                    WriteCrossFormatHeaders(header);
                }
                await using (XlsRowWriter row = await sheet.StartRowAsync(TestContext.Current.CancellationToken))
                {
                    WriteCrossFormatRow(row, value);
                }
                await sheet.EndAsync(TestContext.Current.CancellationToken);
            }
            ms.Position = 0;
            using XlsReader reflectionReader = Excel.FromXls(ms);
            CrossFormatModel reflectionResult = new ExcelParser<CrossFormatModel>().Parse(reflectionReader).First();
            ms.Position = 0;
            using XlsReader generatedReader = Excel.FromXls(ms);
            CrossFormatModel generatedResult = new ExcelMappedParser<CrossFormatModel>().Parse(generatedReader).First();

            AssertCrossFormatEqual(reflectionResult, generatedResult);
        }

        [Fact]
        public async Task GeneratedMapParityAcrossAllFourFormatsCsv()
        {
            var value = SampleCrossFormatValue();
            await using var ms = new MemoryStream();
            CsvWorkbookWriter wb = CsvWorkbookWriter.Create(ms, leaveOpen: true);
            await wb.StartAsync(TestContext.Current.CancellationToken);
            CsvSheetWriter sheet = wb.AddSheet("S1");
            await sheet.StartAsync(TestContext.Current.CancellationToken);
            await using (CsvRowWriter header = await sheet.StartRowAsync(TestContext.Current.CancellationToken))
            {
                WriteCrossFormatHeaders(header);
            }
            await using (CsvRowWriter row = await sheet.StartRowAsync(TestContext.Current.CancellationToken))
            {
                WriteCrossFormatRow(row, value);
            }
            await wb.EndAsync(TestContext.Current.CancellationToken);

            ms.Position = 0;
            CsvReader reflectionReader = Excel.FromCsv(ms);
            CrossFormatModel reflectionResult = new ExcelParser<CrossFormatModel>().Parse(reflectionReader).First();
            ms.Position = 0;
            CsvReader generatedReader = Excel.FromCsv(ms);
            CrossFormatModel generatedResult = new ExcelMappedParser<CrossFormatModel>().Parse(generatedReader).First();

            AssertCrossFormatEqual(reflectionResult, generatedResult);
        }

        private static CrossFormatModel SampleCrossFormatValue()
        {
            return new CrossFormatModel
            {
                Name = "Alice",
                Active = true,
                Age = 30,
                Balance = 12.5m,
                Category = ParityKind.Beta,
                Clock = new TimeOnly(13, 45, 0),
                Id = Guid.Parse("11111111-2222-3333-4444-555555555555"),
                OptionalAge = 7,
            };
        }

        private static void WriteCrossFormatHeaders(IRowWriter row)
        {
            row.Write("Name");
            row.Write("Active");
            row.Write("Age");
            row.Write("Balance");
            row.Write("Category");
            row.Write("Clock");
            row.Write("Id");
            row.Write("OptionalAge");
        }

        private static void WriteCrossFormatRow(IRowWriter row, CrossFormatModel value)
        {
            row.Write(value.Name);
            row.Write(value.Active);
            row.Write(value.Age);
            row.Write(value.Balance);
            row.Write(value.Category.ToString());
            row.Write(value.Clock);
            row.Write(value.Id.ToString());
            row.Write(value.OptionalAge);
        }

        private static void AssertCrossFormatEqual(CrossFormatModel reflectionResult, CrossFormatModel generatedResult)
        {
            Assert.Equal(reflectionResult.Name, generatedResult.Name);
            Assert.Equal(reflectionResult.Active, generatedResult.Active);
            Assert.Equal(reflectionResult.Age, generatedResult.Age);
            Assert.Equal(reflectionResult.Balance, generatedResult.Balance);
            Assert.Equal(reflectionResult.Category, generatedResult.Category);
            Assert.Equal(reflectionResult.Clock, generatedResult.Clock);
            Assert.Equal(reflectionResult.Id, generatedResult.Id);
            Assert.Equal(reflectionResult.OptionalAge, generatedResult.OptionalAge);

            // Pin the actual values, not just mutual agreement.
            Assert.Equal("Alice", generatedResult.Name);
            Assert.Equal(ParityKind.Beta, generatedResult.Category);
        }

        private static async Task<T> ParseFirstXlsxAsync<T>(MemoryStream ms, Func<XlsxReader, IEnumerable<T>> parse)
        {
            ms.Position = 0;
            await using XlsxReader reader = await Excel.FromAsync(ms, ct: TestContext.Current.CancellationToken);
            return parse(reader).First();
        }
    }
}
