using ExcelReader.Core.Parser;
using ExcelReader.Core.Reader;
using ExcelReader.Core.Writer;
using ExcelReader.Core.ValueObjects;

namespace ExcelReader.Tests
{
    // Parity between the reflection path (ExcelParser<T>/WorkbookRecordWriter<TSheet,TRow>) and the
    // source-generated path (ExcelMappedParser<T>/MappedWorkbookRecordWriter<TSheet,TRow>, requiring
    // [ExcelSerializable]) — the generator's own tabela de decisão (type -> reader/writer) is a second,
    // independent copy of ColumnParserFactory's/RecordColumns<T>'s, and the two can diverge silently
    // (see GeneratorTests.InheritedPropertiesAreMapped's history, and GeneratedRecordMapHeadersMatchReflectionHeaders
    // below, which failed before the generator's write side stopped omitting unsupported-type columns).
    // Every model here is declared directly in this file, not as a string literal fed to
    // CSharpGeneratorDriver, so the real build-time analyzer (wired into ExcelReader.Tests.csproj with
    // OutputItemType="Analyzer") produces its IExcelRowMap<T>/IExcelRecordMap<T> for it —
    // ExcelMappedParser<T>/MappedRecordWriter resolve as ordinary compile-time generics, no
    // reflection-over-a-synthetic-assembly gymnastics needed.
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

        // A plain class with no built-in reader and no [ExcelConverter] — reflection's write side
        // (RecordColumns<T>.Plan<TRow>.Build) always writes it via ToString(); before the generator's
        // GetWriteKind stopped defaulting to WriteKind.None, the generated write side silently omitted
        // this column instead. Read side is untouched: neither path can read it back typed, which is
        // fine — this model only exercises the write side (GeneratedRecordMapHeadersMatchReflectionHeaders).
        public sealed class CustomTag
        {
            private readonly string _text;

            public CustomTag(string text)
            {
                _text = text;
            }

            public override string ToString()
            {
                return _text;
            }
        }

        [ExcelSerializable]
        public partial class WriteOnlyTypeModel : InheritedBaseRow
        {
            public string Name { get; set; } = "";
            public CustomTag Tag { get; set; } = new("");
        }

        [ExcelSerializable]
        public partial class RequiredModel
        {
            [ExcelRequired]
            public string Name { get; set; } = "";
        }

        // Includes DateTime/DateOnly: CsvRowWriter writes them as ISO text while XLSX/XLSB/XLS write an
        // Excel serial number — ExcelCellReaders.DateTimeAuto/DateOnlyAuto (T8) is what lets one
        // generated map read both shapes, since ExcelMappedParser<T> reuses a single map across every
        // reader (unlike the reflection path's dedicated csvTextDates map for CSV).
        [ExcelSerializable]
        public partial class CrossFormatModel
        {
            public string Name { get; set; } = "";
            public bool Active { get; set; }
            public int Age { get; set; }
            public decimal Balance { get; set; }
            public ParityKind Category { get; set; }
            public DateTime BirthDate { get; set; }
            public DateOnly BirthDay { get; set; }
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
                BirthDate = new DateTime(2024, 5, 6),
                BirthDay = new DateOnly(2024, 5, 6),
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
            row.Write("BirthDate");
            row.Write("BirthDay");
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
            row.Write(value.BirthDate);
            row.Write(value.BirthDay);
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
            Assert.Equal(reflectionResult.BirthDate, generatedResult.BirthDate);
            Assert.Equal(reflectionResult.BirthDay, generatedResult.BirthDay);
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

        // T10: write-side parity. WorkbookRecordWriter<TSheet,TRow> (reflection, RecordColumns<T>) vs
        // MappedWorkbookRecordWriter<TSheet,TRow> (generated, IExcelRecordMap<T>) writing the same
        // record — read back through the raw, untyped Row/Cell API (not a typed parser) so a column
        // neither path can read back typed (WriteOnlyTypeModel.Tag) still gets compared, text for text.
        private static string[] ReadRowText(Row row)
        {
            var values = new string[row.ColumnCount];
            for (int i = 0; i < row.ColumnCount; i++)
            {
                values[i] = row[i].GetString();
            }
            return values;
        }

        [Fact]
        public async Task GeneratedRecordMapMatchesReflectionOnWriteXlsx()
        {
            CancellationToken ct = TestContext.Current.CancellationToken;
            CrossFormatModel record = SampleCrossFormatValue();

            await using var reflectionMs = new MemoryStream();
            await using (WorkbookRecordWriter<XlsxSheetWriter, XlsxRowWriter> writer = await RecordWriter.CreateXlsxAsync(reflectionMs, leaveOpen: true, ct: ct))
            {
                await writer.WriteSheetAsync("S1", new[] { record }, ct);
            }
            await using var generatedMs = new MemoryStream();
            await using (MappedWorkbookRecordWriter<XlsxSheetWriter, XlsxRowWriter> writer = await MappedRecordWriter.CreateMappedXlsxAsync(generatedMs, leaveOpen: true, ct: ct))
            {
                await writer.WriteSheetAsync("S1", new[] { record }, ct);
            }

            reflectionMs.Position = 0;
            using XlsxReader reflectionReader = Excel.From(reflectionMs);
            using XlsxReader.Enumerator reflectionEnum = reflectionReader.GetEnumerator();
            Assert.True(reflectionEnum.MoveNext());
            string[] reflectionHeaders = ReadRowText(reflectionEnum.Current);
            Assert.True(reflectionEnum.MoveNext());
            string[] reflectionValues = ReadRowText(reflectionEnum.Current);

            generatedMs.Position = 0;
            using XlsxReader generatedReader = Excel.From(generatedMs);
            using XlsxReader.Enumerator generatedEnum = generatedReader.GetEnumerator();
            Assert.True(generatedEnum.MoveNext());
            string[] generatedHeaders = ReadRowText(generatedEnum.Current);
            Assert.True(generatedEnum.MoveNext());
            string[] generatedValues = ReadRowText(generatedEnum.Current);

            Assert.Equal(reflectionHeaders, generatedHeaders);
            Assert.Equal(reflectionValues, generatedValues);
        }

        [Fact]
        public async Task GeneratedRecordMapMatchesReflectionOnWriteXlsb()
        {
            CancellationToken ct = TestContext.Current.CancellationToken;
            CrossFormatModel record = SampleCrossFormatValue();

            await using var reflectionMs = new MemoryStream();
            await using (WorkbookRecordWriter<XlsbSheetWriter, XlsbRowWriter> writer = await RecordWriter.CreateXlsbAsync(reflectionMs, leaveOpen: true, ct: ct))
            {
                await writer.WriteSheetAsync("S1", new[] { record }, ct);
            }
            await using var generatedMs = new MemoryStream();
            await using (MappedWorkbookRecordWriter<XlsbSheetWriter, XlsbRowWriter> writer = await MappedRecordWriter.CreateMappedXlsbAsync(generatedMs, leaveOpen: true, ct: ct))
            {
                await writer.WriteSheetAsync("S1", new[] { record }, ct);
            }

            reflectionMs.Position = 0;
            using XlsbReader reflectionReader = Excel.FromXlsb(reflectionMs);
            using XlsbReader.Enumerator reflectionEnum = reflectionReader.GetEnumerator();
            Assert.True(reflectionEnum.MoveNext());
            string[] reflectionHeaders = ReadRowText(reflectionEnum.Current);
            Assert.True(reflectionEnum.MoveNext());
            string[] reflectionValues = ReadRowText(reflectionEnum.Current);

            generatedMs.Position = 0;
            using XlsbReader generatedReader = Excel.FromXlsb(generatedMs);
            using XlsbReader.Enumerator generatedEnum = generatedReader.GetEnumerator();
            Assert.True(generatedEnum.MoveNext());
            string[] generatedHeaders = ReadRowText(generatedEnum.Current);
            Assert.True(generatedEnum.MoveNext());
            string[] generatedValues = ReadRowText(generatedEnum.Current);

            Assert.Equal(reflectionHeaders, generatedHeaders);
            Assert.Equal(reflectionValues, generatedValues);
        }

        [Fact]
        public async Task GeneratedRecordMapMatchesReflectionOnWriteXls()
        {
            CancellationToken ct = TestContext.Current.CancellationToken;
            CrossFormatModel record = SampleCrossFormatValue();

            await using var reflectionMs = new MemoryStream();
            await using (WorkbookRecordWriter<XlsSheetWriter, XlsRowWriter> writer = await RecordWriter.CreateXlsAsync(reflectionMs, leaveOpen: true, ct: ct))
            {
                await writer.WriteSheetAsync("S1", new[] { record }, ct);
            }
            await using var generatedMs = new MemoryStream();
            await using (MappedWorkbookRecordWriter<XlsSheetWriter, XlsRowWriter> writer = await MappedRecordWriter.CreateMappedXlsAsync(generatedMs, leaveOpen: true, ct: ct))
            {
                await writer.WriteSheetAsync("S1", new[] { record }, ct);
            }

            reflectionMs.Position = 0;
            using XlsReader reflectionReader = Excel.FromXls(reflectionMs);
            using XlsReader.Enumerator reflectionEnum = reflectionReader.GetEnumerator();
            Assert.True(reflectionEnum.MoveNext());
            string[] reflectionHeaders = ReadRowText(reflectionEnum.Current);
            Assert.True(reflectionEnum.MoveNext());
            string[] reflectionValues = ReadRowText(reflectionEnum.Current);

            generatedMs.Position = 0;
            using XlsReader generatedReader = Excel.FromXls(generatedMs);
            using XlsReader.Enumerator generatedEnum = generatedReader.GetEnumerator();
            Assert.True(generatedEnum.MoveNext());
            string[] generatedHeaders = ReadRowText(generatedEnum.Current);
            Assert.True(generatedEnum.MoveNext());
            string[] generatedValues = ReadRowText(generatedEnum.Current);

            Assert.Equal(reflectionHeaders, generatedHeaders);
            Assert.Equal(reflectionValues, generatedValues);
        }

        [Fact]
        public async Task GeneratedRecordMapMatchesReflectionOnWriteCsv()
        {
            CancellationToken ct = TestContext.Current.CancellationToken;
            CrossFormatModel record = SampleCrossFormatValue();

            await using var reflectionMs = new MemoryStream();
            await using (WorkbookRecordWriter<CsvSheetWriter, CsvRowWriter> writer = await RecordWriter.CreateCsvAsync(reflectionMs, leaveOpen: true, ct: ct))
            {
                await writer.WriteSheetAsync("S1", new[] { record }, ct);
            }
            await using var generatedMs = new MemoryStream();
            await using (MappedWorkbookRecordWriter<CsvSheetWriter, CsvRowWriter> writer = await MappedRecordWriter.CreateMappedCsvAsync(generatedMs, leaveOpen: true, ct: ct))
            {
                await writer.WriteSheetAsync("S1", new[] { record }, ct);
            }

            reflectionMs.Position = 0;
            using CsvReader reflectionReader = Excel.FromCsv(reflectionMs);
            using CsvReader.Enumerator reflectionEnum = reflectionReader.GetEnumerator();
            Assert.True(reflectionEnum.MoveNext());
            string[] reflectionHeaders = ReadRowText(reflectionEnum.Current);
            Assert.True(reflectionEnum.MoveNext());
            string[] reflectionValues = ReadRowText(reflectionEnum.Current);

            generatedMs.Position = 0;
            using CsvReader generatedReader = Excel.FromCsv(generatedMs);
            using CsvReader.Enumerator generatedEnum = generatedReader.GetEnumerator();
            Assert.True(generatedEnum.MoveNext());
            string[] generatedHeaders = ReadRowText(generatedEnum.Current);
            Assert.True(generatedEnum.MoveNext());
            string[] generatedValues = ReadRowText(generatedEnum.Current);

            Assert.Equal(reflectionHeaders, generatedHeaders);
            Assert.Equal(reflectionValues, generatedValues);
        }

        [Fact]
        public async Task GeneratedRecordMapHeadersMatchReflectionHeaders()
        {
            CancellationToken ct = TestContext.Current.CancellationToken;
            var record = new WriteOnlyTypeModel { Inherited = "BaseValue", Name = "Alice", Tag = new CustomTag("T1") };

            await using var reflectionMs = new MemoryStream();
            await using (WorkbookRecordWriter<XlsxSheetWriter, XlsxRowWriter> writer = await RecordWriter.CreateXlsxAsync(reflectionMs, leaveOpen: true, ct: ct))
            {
                await writer.WriteSheetAsync("S1", new[] { record }, ct);
            }
            await using var generatedMs = new MemoryStream();
            await using (MappedWorkbookRecordWriter<XlsxSheetWriter, XlsxRowWriter> writer = await MappedRecordWriter.CreateMappedXlsxAsync(generatedMs, leaveOpen: true, ct: ct))
            {
                await writer.WriteSheetAsync("S1", new[] { record }, ct);
            }

            reflectionMs.Position = 0;
            using XlsxReader reflectionReader = Excel.From(reflectionMs);
            using XlsxReader.Enumerator reflectionEnum = reflectionReader.GetEnumerator();
            Assert.True(reflectionEnum.MoveNext());
            string[] reflectionHeaders = ReadRowText(reflectionEnum.Current);
            Assert.True(reflectionEnum.MoveNext());
            string[] reflectionValues = ReadRowText(reflectionEnum.Current);

            generatedMs.Position = 0;
            using XlsxReader generatedReader = Excel.From(generatedMs);
            using XlsxReader.Enumerator generatedEnum = generatedReader.GetEnumerator();
            Assert.True(generatedEnum.MoveNext());
            string[] generatedHeaders = ReadRowText(generatedEnum.Current);
            Assert.True(generatedEnum.MoveNext());
            string[] generatedValues = ReadRowText(generatedEnum.Current);

            // Before the generator's GetWriteKind stopped defaulting to WriteKind.None for an
            // unrecognized type, "Tag" would be missing entirely from generatedHeaders/generatedValues
            // here, and the header-count mismatch alone would fail this assertion.
            Assert.Equal(reflectionHeaders, generatedHeaders);
            Assert.Equal(reflectionValues, generatedValues);
            Assert.Contains("Tag", generatedHeaders, StringComparer.Ordinal);
            Assert.Contains("T1", generatedValues, StringComparer.Ordinal);
        }
    }
}
