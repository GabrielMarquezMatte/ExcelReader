using ExcelReader.Core.Parser;
using ExcelReader.Core.Parser.Internal;
using ExcelReader.Core.Reader;
using ExcelReader.Core.Writer;

namespace ExcelReader.Tests
{
    // Feature A0 (docs/v2-plan.md §2.4): proves the hand-written ExcelRowMapBuilder<T>/ExcelRecordMapBuilder<T>
    // seam produces the exact same result as the reflection-based ExcelParser<T>/WorkbookRecordWriter for the
    // same model, before any source generator exists to emit the map automatically.
    public class ExcelRowMapBuilderTests
    {
        private enum MapBuilderKind
        {
            Alpha,
            Beta,
        }

        private sealed class MapBuilderTestModel : IExcelRowMap<MapBuilderTestModel>, IExcelRecordMap<MapBuilderTestModel>
        {
            public string Name { get; set; } = "";
            public int Age { get; set; }
            public DateTime BirthDate { get; set; }
            public bool Active { get; set; }
            public MapBuilderKind Kind { get; set; }

            public static void ConfigureExcelRowMap(ExcelRowMapBuilder<MapBuilderTestModel> builder)
            {
                builder
                    .Factory(static () => new MapBuilderTestModel())
                    .Property(["Name"], ExcelCellReaders.String, static (ref m, v) => m.Name = v)
                    .Property(["Age"], ExcelCellReaders.Parsable, static (ref MapBuilderTestModel m, int v) => m.Age = v)
                    .Property(["BirthDate"], ExcelCellReaders.DateTimeSerial, static (ref m, v) => m.BirthDate = v)
                    .Property(["Active"], ExcelCellReaders.Bool, static (ref m, v) => m.Active = v)
                    .Property(["Kind"], ExcelCellReaders.Enum, static (ref MapBuilderTestModel m, MapBuilderKind v) => m.Kind = v);
            }

            public static void ConfigureExcelRecordMap(ExcelRecordMapBuilder<MapBuilderTestModel> builder)
            {
                builder
                    .Column("Name", static (row, m) => row.Write(m.Name))
                    .Column("Age", static (row, m) => row.Write(m.Age))
                    .Column("BirthDate", static (row, m) => row.Write(m.BirthDate))
                    .Column("Active", static (row, m) => row.Write(m.Active))
                    .Column("Kind", static (row, m) => row.Write(m.Kind.ToString()));
            }
        }

        // Mirrors MapBuilderTestModel's shape exactly, so ExcelParser<T>'s reflection path binds the same
        // headers the same way, giving a same-source-model reflection baseline to compare against.
        private sealed class ReflectionTestModel
        {
            public string Name { get; set; } = "";
            public int Age { get; set; }
            public DateTime BirthDate { get; set; }
            public bool Active { get; set; }
            public MapBuilderKind Kind { get; set; }
        }

        private static readonly DateTime SampleBirthDate = new(1990, 5, 17);

        [Fact]
        public async Task MappedParserMatchesReflectionOnRead()
        {
            await using MemoryStream ms = await TypedWorkbook.BuildAsync(
                ["Name", "Age", "BirthDate", "Active", "Kind"],
                ["Alice", 30, SampleBirthDate, true, "Beta"]);

            await using XlsxReader mappedReader = await Excel.FromAsync(ms, leaveOpen: true, ct: TestContext.Current.CancellationToken);
            List<MapBuilderTestModel> mapped = new ExcelMappedParser<MapBuilderTestModel>().Parse(mappedReader).ToList();

            ms.Position = 0;
            await using XlsxReader reflectionReader = await Excel.FromAsync(ms, ct: TestContext.Current.CancellationToken);
            List<ReflectionTestModel> reflected = new ExcelParser<ReflectionTestModel>().Parse(reflectionReader).ToList();

            Assert.Single(mapped);
            Assert.Single(reflected);
            Assert.Equal(reflected[0].Name, mapped[0].Name);
            Assert.Equal(reflected[0].Age, mapped[0].Age);
            Assert.Equal(reflected[0].BirthDate, mapped[0].BirthDate);
            Assert.Equal(reflected[0].Active, mapped[0].Active);
            Assert.Equal(reflected[0].Kind, mapped[0].Kind);
        }

        [Fact]
        public async Task RequiredColumnMissingThrowsLikeReflection()
        {
            await using MemoryStream mappedStream = await TypedWorkbook.BuildAsync(["Name"], ["Alice"]);
            var builder = new ExcelRowMapBuilder<MapBuilderTestModel>();
            builder
                .Factory(static () => new MapBuilderTestModel())
                .Property(["Age"], ExcelCellReaders.Parsable, static (ref MapBuilderTestModel m, int v) => m.Age = v, isRequired: true);
            TypeMapInfo<MapBuilderTestModel> info = builder.Build();

            await using XlsxReader mappedReader = await Excel.FromAsync(mappedStream, ct: TestContext.Current.CancellationToken);
            var mappedEnumerable = new ExcelEnumerable<MapBuilderTestModel>(mappedReader, new ExcelParserConfig(), info);
            ExcelParseException mappedEx = Assert.Throws<ExcelParseException>(() => mappedEnumerable.ToList());

            await using MemoryStream reflectedStream = await TypedWorkbook.BuildAsync(["Name"], ["Alice"]);
            await using XlsxReader reflectedReader = await Excel.FromAsync(reflectedStream, ct: TestContext.Current.CancellationToken);
            ExcelParseException reflectedEx = Assert.Throws<ExcelParseException>(
                () => new ExcelParser<RequiredAgeRow>().Parse(reflectedReader).ToList());

            Assert.Equal(reflectedEx.Message, mappedEx.Message);
        }

        private sealed class RequiredAgeRow
        {
            public string? Name { get; set; }
            [ExcelRequired]
            public int Age { get; set; }
        }

        [Fact]
        public async Task RecordMapBuilderMatchesReflectionOnWrite()
        {
            var record = new MapBuilderTestModel { Name = "Bob", Age = 42, BirthDate = SampleBirthDate, Active = true, Kind = MapBuilderKind.Alpha };

            CancellationToken ct = TestContext.Current.CancellationToken;
            await using var mappedStream = new MemoryStream();
            await using (XlsxWorkbookWriter wb = await XlsxWorkbookWriter.CreateAsync(mappedStream, leaveOpen: true, ct: ct))
            {
                await wb.StartAsync(ct);
                XlsxSheetWriter sheet = wb.AddSheet("S1");
                await sheet.StartAsync(ct);
                var builder = new ExcelRecordMapBuilder<MapBuilderTestModel>();
                MapBuilderTestModel.ConfigureExcelRecordMap(builder);
                await WriteMappedRowAsync(sheet, builder, record, ct);
                await sheet.EndAsync(ct);
                await wb.EndAsync(ct);
            }

            await using var reflectedStream = new MemoryStream();
            await using (WorkbookRecordWriter<XlsxSheetWriter, XlsxRowWriter> writer =
                await RecordWriter.CreateXlsxAsync(reflectedStream, leaveOpen: true, ct: ct))
            {
                await writer.WriteSheetAsync("S1", [record], ct);
            }

            mappedStream.Position = 0;
            reflectedStream.Position = 0;
            await using XlsxReader mappedReader = await Excel.FromAsync(mappedStream, ct: TestContext.Current.CancellationToken);
            await using XlsxReader reflectedReader = await Excel.FromAsync(reflectedStream, ct: TestContext.Current.CancellationToken);
            List<ReflectionTestModel> mapped = new ExcelParser<ReflectionTestModel>().Parse(mappedReader).ToList();
            List<ReflectionTestModel> reflected = new ExcelParser<ReflectionTestModel>().Parse(reflectedReader).ToList();

            Assert.Single(mapped);
            Assert.Single(reflected);
            Assert.Equal(reflected[0].Name, mapped[0].Name);
            Assert.Equal(reflected[0].Age, mapped[0].Age);
            Assert.Equal(reflected[0].BirthDate, mapped[0].BirthDate);
            Assert.Equal(reflected[0].Active, mapped[0].Active);
            Assert.Equal(reflected[0].Kind, mapped[0].Kind);
        }

        private static async ValueTask WriteMappedRowAsync(XlsxSheetWriter sheet, ExcelRecordMapBuilder<MapBuilderTestModel> builder, MapBuilderTestModel record, CancellationToken ct)
        {
            XlsxRowWriter header = await sheet.StartRowAsync(ct);
            await using (header.ConfigureAwait(false))
            {
                foreach (string h in builder.Headers())
                {
                    header.Write(h);
                }
            }
            XlsxRowWriter row = await sheet.StartRowAsync(ct);
            await using (row.ConfigureAwait(false))
            {
                builder.WriteRow(row, record);
            }
        }
    }
}
