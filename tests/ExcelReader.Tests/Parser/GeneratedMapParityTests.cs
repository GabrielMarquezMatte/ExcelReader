using ExcelReader.Core.Parser;
using ExcelReader.Core.Parser.Internal;
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

        public class InitBaseRow
        {
            public string BaseInit { get; init; } = "";
        }

        [ExcelSerializable]
        public partial class SetterShapesModel : InitBaseRow
        {
            public string Settable { get; set; } = "";
            public string InitOnly { get; init; } = "";
            public int? NullableInit { get; init; }
            [ExcelConverter(typeof(UpperCaseConverter))]
            public string ConvertedInit { get; init; } = "";
            public string GetOnly { get; } = "default";
            public string PrivateSet { get; private set; } = "default";
            private string Hidden { get; set; } = "default";
            public string HiddenValue => Hidden;
        }

#pragma warning disable CA1815
        [ExcelSerializable]
        public partial struct InitStructModel
        {
            public string Name { get; init; }
            public int Age { get; init; }
        }
#pragma warning restore CA1815

        [ExcelSerializable]
        public partial record InitRecordModel
        {
            public string Name { get; init; } = "";
        }

        [ExcelSerializable]
        public readonly ref partial struct InitRefStructModel
        {
            public ReadOnlySpan<byte> Name { get; init; }
            public int Age { get; init; }
        }

        [ExcelSerializable]
        public partial class EscapedHeaderModel
        {
            [ExcelColumn("say \"hi\"")]
            public string Quote { get; set; } = "";
            [ExcelColumn(@"C:\temp\new")]
            public string Backslash { get; set; } = "";
            [ExcelColumn("a\tb")]
            public string Tab { get; set; } = "";
            [ExcelColumn("a\rb")]
            public string CarriageReturn { get; set; } = "";
            [ExcelColumn("a\nb")]
            public string LineFeed { get; set; } = "";
            [ExcelColumn("")]
            public string Empty { get; set; } = "";
            [ExcelColumn("Ação ✓ 😀 日本")]
            public string Unicode { get; set; } = "";
            [ExcelColumn("nul\0 bel\a esc\u001B del\u007F nel\u0085 ls\u2028 ps\u2029 bom\uFEFF")]
            public string Control { get; set; } = "";
            [ExcelColumn("primary {0} $\"")]
            [ExcelColumn("alias\\n \"x\"\r\n")]
            public string Aliased { get; set; } = "";
        }

        private static readonly string[] EscapedHeaders =
        [
            "say \"hi\"", @"C:\temp\new", "a\tb", "a\rb", "a\nb", "",
            "Ação ✓ 😀 日本", "nul\0 bel\a esc\u001B del\u007F nel\u0085 ls\u2028 ps\u2029 bom\uFEFF", "primary {0} $\"",
        ];

        [Fact]
        public async Task GeneratedMapMatchesReflectionForEverySetterShape()
        {
            await using var ms = await TypedWorkbook.BuildAsync(
                ["BaseInit", "Settable", "InitOnly", "NullableInit", "ConvertedInit", "GetOnly", "PrivateSet", "Hidden"],
                ["base", "set", "init", 7, "conv", "x", "y", "z"]);

            SetterShapesModel reflectionResult = await ParseFirstXlsxAsync(ms, static reader => ExcelParser.FromAttributes<SetterShapesModel>().Parse(reader));
            SetterShapesModel generatedResult = await ParseFirstXlsxAsync(ms, static reader => ExcelParser.Generated<SetterShapesModel>().Parse(reader));

            foreach (SetterShapesModel result in new[] { reflectionResult, generatedResult })
            {
                Assert.Equal("base", result.BaseInit);
                Assert.Equal("set", result.Settable);
                Assert.Equal("init", result.InitOnly);
                Assert.Equal(7, result.NullableInit);
                Assert.Equal("CONV", result.ConvertedInit);
                Assert.Equal("default", result.GetOnly);
                Assert.Equal("default", result.PrivateSet);
                Assert.Equal("default", result.HiddenValue);
            }
        }

        [Fact]
        public async Task GeneratedMapMatchesReflectionForInitOnlyStructAndRecord()
        {
            await using var ms = await TypedWorkbook.BuildAsync(["Name", "Age"], ["Alice", 30]);

            InitStructModel reflectionStruct = await ParseFirstXlsxAsync(ms, static reader => ExcelParser.FromAttributes<InitStructModel>().Parse(reader));
            InitStructModel generatedStruct = await ParseFirstXlsxAsync(ms, static reader => ExcelParser.Generated<InitStructModel>().Parse(reader));
            InitRecordModel reflectionRecord = await ParseFirstXlsxAsync(ms, static reader => ExcelParser.FromAttributes<InitRecordModel>().Parse(reader));
            InitRecordModel generatedRecord = await ParseFirstXlsxAsync(ms, static reader => ExcelParser.Generated<InitRecordModel>().Parse(reader));

            Assert.Equal(new InitStructModel { Name = "Alice", Age = 30 }, reflectionStruct);
            Assert.Equal(reflectionStruct, generatedStruct);
            Assert.Equal("Alice", reflectionRecord.Name);
            Assert.Equal(reflectionRecord, generatedRecord);
        }

        [Fact]
        public async Task GeneratedMapMatchesReflectionForInitOnlyRefStruct()
        {
            await using var ms = await TypedWorkbook.BuildAsync(["Name", "Age"], ["Alice", 30]);

            using var reflectionReader = Excel.FromXlsx(ms, leaveOpen: true);
            var reflectionEnum = ExcelParser.FromAttributes<InitRefStructModel>().Parse(reflectionReader).GetEnumerator();
            Assert.True(reflectionEnum.MoveNext());
            InitRefStructModel reflectionRow = reflectionEnum.Current;
            (string Name, int Age) reflectionResult = (System.Text.Encoding.UTF8.GetString(reflectionRow.Name), reflectionRow.Age);

            ms.Position = 0;
            using var generatedReader = Excel.FromXlsx(ms, leaveOpen: true);
            var generatedEnum = ExcelParser.Generated<InitRefStructModel>().Parse(generatedReader).GetEnumerator();
            Assert.True(generatedEnum.MoveNext());
            InitRefStructModel generatedRow = generatedEnum.Current;
            (string Name, int Age) generatedResult = (System.Text.Encoding.UTF8.GetString(generatedRow.Name), generatedRow.Age);

            Assert.Equal(("Alice", 30), reflectionResult);
            Assert.Equal(reflectionResult, generatedResult);
        }

        [Fact]
        public void GeneratedReadMapKeepsEscapedHeaderTextExact()
        {
            var builder = new ExcelRowMapBuilder<EscapedHeaderModel>();
            EscapedHeaderModel.ConfigureExcelRowMap(builder);
            TypeMapInfo<EscapedHeaderModel> generated = builder.Build();
            TypeMapInfo<EscapedHeaderModel> reflection = TypeMapper<EscapedHeaderModel>.GetInfo();

            Assert.Equal(reflection.PropertyCount, generated.PropertyCount);
            foreach (string header in (string[])[.. EscapedHeaders, "alias\\n \"x\"\r\n"])
            {
                Assert.True(generated.TryFindHeader(header, StringComparer.Ordinal, HeaderNormalization.None, out var generatedMatch), header);
                Assert.True(reflection.TryFindHeader(header, StringComparer.Ordinal, HeaderNormalization.None, out var reflectionMatch), header);
                Assert.Equal(reflectionMatch.AliasIndex, generatedMatch.AliasIndex);
                Assert.Equal(reflection.DisplayName(reflectionMatch.PropertyIndex), generated.DisplayName(generatedMatch.PropertyIndex));
            }
            Assert.Equal(EscapedHeaders, Enumerable.Range(0, generated.PropertyCount).Select(generated.DisplayName), StringComparer.Ordinal);
        }

        [Fact]
        public void GeneratedRecordMapKeepsEscapedHeaderTextExact()
        {
            string[] generated = ExcelRecordLayout.Generated<EscapedHeaderModel>().Columns<CsvRowWriter>().Headers;
            string[] reflection = ExcelRecordLayout.FromAttributes<EscapedHeaderModel>().Columns<CsvRowWriter>().Headers;

            Assert.Equal(EscapedHeaders, generated);
            Assert.Equal(reflection, generated);
        }

        [Fact]
        public async Task GeneratedMapMatchesReflectionForEveryBuiltInType()
        {
            var kind = ParityKind.Beta;
            var id = Guid.Parse("11111111-2222-3333-4444-555555555555");
            object?[] row =
            [
                "Alice", true, new DateTime(2024, 5, 6), new DateTime(2024, 5, 6),
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

            EveryTypeModel reflectionResult = await ParseFirstXlsxAsync(ms, static reader => ExcelParser.FromAttributes<EveryTypeModel>().Parse(reader));
            EveryTypeModel generatedResult = await ParseFirstXlsxAsync(ms, static reader => ExcelParser.Generated<EveryTypeModel>().Parse(reader));

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

            Assert.Equal("Alice", generatedResult.Name);
            Assert.Equal(ParityKind.Beta, generatedResult.Category);
            Assert.Equal(id, generatedResult.Id);
        }

        [Fact]
        public async Task GeneratedMapMatchesReflectionForCustomConverter()
        {
            await using var ms = await TypedWorkbook.BuildAsync(["Name"], ["alice"]);

            ConverterModel reflectionResult = await ParseFirstXlsxAsync(ms, static reader => ExcelParser.FromAttributes<ConverterModel>().Parse(reader));
            ConverterModel generatedResult = await ParseFirstXlsxAsync(ms, static reader => ExcelParser.Generated<ConverterModel>().Parse(reader));

            Assert.Equal("ALICE", reflectionResult.Name);
            Assert.Equal("ALICE", generatedResult.Name);
        }

        [Fact]
        public async Task GeneratedMapHonorsColumnAliases()
        {
            await using var ms = await TypedWorkbook.BuildAsync(["Name"], ["Alice"]);

            AliasModel reflectionResult = await ParseFirstXlsxAsync(ms, static reader => ExcelParser.FromAttributes<AliasModel>().Parse(reader));
            AliasModel generatedResult = await ParseFirstXlsxAsync(ms, static reader => ExcelParser.Generated<AliasModel>().Parse(reader));

            Assert.Equal("Alice", reflectionResult.Name);
            Assert.Equal("Alice", generatedResult.Name);
        }

        [Fact]
        public async Task GeneratedMapHonorsIgnore()
        {
            await using var ms = await TypedWorkbook.BuildAsync(["Name", "Ignored"], ["Alice", "SHOULD_NOT_BIND"]);

            IgnoreModel reflectionResult = await ParseFirstXlsxAsync(ms, static reader => ExcelParser.FromAttributes<IgnoreModel>().Parse(reader));
            IgnoreModel generatedResult = await ParseFirstXlsxAsync(ms, static reader => ExcelParser.Generated<IgnoreModel>().Parse(reader));

            Assert.Equal("Alice", reflectionResult.Name);
            Assert.Equal("default", reflectionResult.Ignored);
            Assert.Equal("Alice", generatedResult.Name);
            Assert.Equal("default", generatedResult.Ignored);
        }

        [Fact]
        public async Task GeneratedMapMatchesReflectionForInheritedProperties()
        {
            await using var ms = await TypedWorkbook.BuildAsync(["Inherited", "Own"], ["BaseValue", 42]);

            InheritedModel reflectionResult = await ParseFirstXlsxAsync(ms, static reader => ExcelParser.FromAttributes<InheritedModel>().Parse(reader));
            InheritedModel generatedResult = await ParseFirstXlsxAsync(ms, static reader => ExcelParser.Generated<InheritedModel>().Parse(reader));

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
                () => ParseFirstXlsxAsync(ms, static reader => ExcelParser.FromAttributes<RequiredModel>().Parse(reader)));
            ExcelParseException generatedException = await Assert.ThrowsAsync<ExcelParseException>(
                () => ParseFirstXlsxAsync(ms, static reader => ExcelParser.Generated<RequiredModel>().Parse(reader)));

            Assert.Contains("Name", reflectionException.ColumnName, StringComparison.Ordinal);
            Assert.Contains("Name", generatedException.ColumnName, StringComparison.Ordinal);
        }

        [Fact]
        public async Task GeneratedMapParityAcrossAllFourFormatsXlsx()
        {
            var value = SampleCrossFormatValue();
            await using var ms = new MemoryStream();
            await using (XlsxWorkbookWriter wb = XlsxWorkbookWriter.Create(ms, leaveOpen: true))
            {
                XlsxSheetWriter sheet = wb.AddSheet("S1");
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
            await using XlsxReader reflectionReader = await Excel.FromXlsxAsync(ms, ct: TestContext.Current.CancellationToken);
            CrossFormatModel reflectionResult = ExcelParser.FromAttributes<CrossFormatModel>().Parse(reflectionReader).First();
            ms.Position = 0;
            await using XlsxReader generatedReader = await Excel.FromXlsxAsync(ms, ct: TestContext.Current.CancellationToken);
            CrossFormatModel generatedResult = ExcelParser.Generated<CrossFormatModel>().Parse(generatedReader).First();

            AssertCrossFormatEqual(reflectionResult, generatedResult);
        }

        [Fact]
        public async Task GeneratedMapParityAcrossAllFourFormatsXlsb()
        {
            var value = SampleCrossFormatValue();
            await using var ms = new MemoryStream();
            await using (XlsbWorkbookWriter wb = XlsbWorkbookWriter.Create(ms, leaveOpen: true))
            {
                XlsbSheetWriter sheet = wb.AddSheet("S1");
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
            CrossFormatModel reflectionResult = ExcelParser.FromAttributes<CrossFormatModel>().Parse(reflectionReader).First();
            ms.Position = 0;
            using XlsbReader generatedReader = Excel.FromXlsb(ms);
            CrossFormatModel generatedResult = ExcelParser.Generated<CrossFormatModel>().Parse(generatedReader).First();

            AssertCrossFormatEqual(reflectionResult, generatedResult);
        }

        [Fact]
        public async Task GeneratedMapParityAcrossAllFourFormatsXls()
        {
            var value = SampleCrossFormatValue();
            await using var ms = new MemoryStream();
            await using (XlsWorkbookWriter wb = XlsWorkbookWriter.Create(ms, leaveOpen: true))
            {
                XlsSheetWriter sheet = wb.AddSheet("S1");
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
            CrossFormatModel reflectionResult = ExcelParser.FromAttributes<CrossFormatModel>().Parse(reflectionReader).First();
            ms.Position = 0;
            using XlsReader generatedReader = Excel.FromXls(ms);
            CrossFormatModel generatedResult = ExcelParser.Generated<CrossFormatModel>().Parse(generatedReader).First();

            AssertCrossFormatEqual(reflectionResult, generatedResult);
        }

        [Fact]
        public async Task GeneratedMapParityAcrossAllFourFormatsCsv()
        {
            var value = SampleCrossFormatValue();
            await using var ms = new MemoryStream();
            CsvWorkbookWriter wb = CsvWorkbookWriter.Create(ms, leaveOpen: true);
            CsvSheetWriter sheet = wb.AddSheet("S1");
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
            CrossFormatModel reflectionResult = ExcelParser.FromAttributes<CrossFormatModel>().Parse(reflectionReader).First();
            ms.Position = 0;
            CsvReader generatedReader = Excel.FromCsv(ms);
            CrossFormatModel generatedResult = ExcelParser.Generated<CrossFormatModel>().Parse(generatedReader).First();

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

            Assert.Equal("Alice", generatedResult.Name);
            Assert.Equal(ParityKind.Beta, generatedResult.Category);
        }

        private static async Task<T> ParseFirstXlsxAsync<T>(MemoryStream ms, Func<XlsxReader, IEnumerable<T>> parse)
        {
            ms.Position = 0;
            await using XlsxReader reader = await Excel.FromXlsxAsync(ms, ct: TestContext.Current.CancellationToken);
            return parse(reader).First();
        }

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
            await using (XlsxWorkbookWriter writer = XlsxWorkbookWriter.Create(reflectionMs, leaveOpen: true))
            {
                await using (var sheet = writer.AddSheet("S1"))
                {
                    await sheet.WriteRecordsAsync(new[] { record }, ExcelRecordLayout.FromAttributes<CrossFormatModel>(), ct);
                }
            }
            await using var generatedMs = new MemoryStream();
            await using (XlsxWorkbookWriter writer = XlsxWorkbookWriter.Create(generatedMs, leaveOpen: true))
            {
                await using (var sheet = writer.AddSheet("S1"))
                {
                    await sheet.WriteRecordsAsync(new[] { record }, ExcelRecordLayout.Generated<CrossFormatModel>(), ct);
                }
            }

            reflectionMs.Position = 0;
            using XlsxReader reflectionReader = Excel.FromXlsx(reflectionMs);
            using XlsxReader.Enumerator reflectionEnum = reflectionReader.GetEnumerator();
            Assert.True(reflectionEnum.MoveNext());
            string[] reflectionHeaders = ReadRowText(reflectionEnum.Current);
            Assert.True(reflectionEnum.MoveNext());
            string[] reflectionValues = ReadRowText(reflectionEnum.Current);

            generatedMs.Position = 0;
            using XlsxReader generatedReader = Excel.FromXlsx(generatedMs);
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
            await using (XlsbWorkbookWriter writer = XlsbWorkbookWriter.Create(reflectionMs, leaveOpen: true))
            {
                await using (var sheet = writer.AddSheet("S1"))
                {
                    await sheet.WriteRecordsAsync(new[] { record }, ExcelRecordLayout.FromAttributes<CrossFormatModel>(), ct);
                }
            }
            await using var generatedMs = new MemoryStream();
            await using (XlsbWorkbookWriter writer = XlsbWorkbookWriter.Create(generatedMs, leaveOpen: true))
            {
                await using (var sheet = writer.AddSheet("S1"))
                {
                    await sheet.WriteRecordsAsync(new[] { record }, ExcelRecordLayout.Generated<CrossFormatModel>(), ct);
                }
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
            await using (XlsWorkbookWriter writer = XlsWorkbookWriter.Create(reflectionMs, leaveOpen: true))
            {
                await using (var sheet = writer.AddSheet("S1"))
                {
                    await sheet.WriteRecordsAsync(new[] { record }, ExcelRecordLayout.FromAttributes<CrossFormatModel>(), ct);
                }
            }
            await using var generatedMs = new MemoryStream();
            await using (XlsWorkbookWriter writer = XlsWorkbookWriter.Create(generatedMs, leaveOpen: true))
            {
                await using (var sheet = writer.AddSheet("S1"))
                {
                    await sheet.WriteRecordsAsync(new[] { record }, ExcelRecordLayout.Generated<CrossFormatModel>(), ct);
                }
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
            await using (CsvWorkbookWriter writer = CsvWorkbookWriter.Create(reflectionMs, leaveOpen: true))
            {
                await using (var sheet = writer.AddSheet("S1"))
                {
                    await sheet.WriteRecordsAsync(new[] { record }, ExcelRecordLayout.FromAttributes<CrossFormatModel>(), ct);
                }
            }
            await using var generatedMs = new MemoryStream();
            await using (CsvWorkbookWriter writer = CsvWorkbookWriter.Create(generatedMs, leaveOpen: true))
            {
                await using (var sheet = writer.AddSheet("S1"))
                {
                    await sheet.WriteRecordsAsync(new[] { record }, ExcelRecordLayout.Generated<CrossFormatModel>(), ct);
                }
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
            await using (XlsxWorkbookWriter writer = XlsxWorkbookWriter.Create(reflectionMs, leaveOpen: true))
            {
                await using (var sheet = writer.AddSheet("S1"))
                {
                    await sheet.WriteRecordsAsync(new[] { record }, ExcelRecordLayout.FromAttributes<WriteOnlyTypeModel>(), ct);
                }
            }
            await using var generatedMs = new MemoryStream();
            await using (XlsxWorkbookWriter writer = XlsxWorkbookWriter.Create(generatedMs, leaveOpen: true))
            {
                await using (var sheet = writer.AddSheet("S1"))
                {
                    await sheet.WriteRecordsAsync(new[] { record }, ExcelRecordLayout.Generated<WriteOnlyTypeModel>(), ct);
                }
            }

            reflectionMs.Position = 0;
            using XlsxReader reflectionReader = Excel.FromXlsx(reflectionMs);
            using XlsxReader.Enumerator reflectionEnum = reflectionReader.GetEnumerator();
            Assert.True(reflectionEnum.MoveNext());
            string[] reflectionHeaders = ReadRowText(reflectionEnum.Current);
            Assert.True(reflectionEnum.MoveNext());
            string[] reflectionValues = ReadRowText(reflectionEnum.Current);

            generatedMs.Position = 0;
            using XlsxReader generatedReader = Excel.FromXlsx(generatedMs);
            using XlsxReader.Enumerator generatedEnum = generatedReader.GetEnumerator();
            Assert.True(generatedEnum.MoveNext());
            string[] generatedHeaders = ReadRowText(generatedEnum.Current);
            Assert.True(generatedEnum.MoveNext());
            string[] generatedValues = ReadRowText(generatedEnum.Current);

            Assert.Equal(reflectionHeaders, generatedHeaders);
            Assert.Equal(reflectionValues, generatedValues);
            Assert.Contains("Tag", generatedHeaders, StringComparer.Ordinal);
            Assert.Contains("T1", generatedValues, StringComparer.Ordinal);
        }
    }
}
