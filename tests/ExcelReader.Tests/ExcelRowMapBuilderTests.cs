using ExcelReader.Core.Parser;
using ExcelReader.Core.Parser.Internal;
using ExcelReader.Core.Reader;
using ExcelReader.Core.Writer;

namespace ExcelReader.Tests
{
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

            public static void ConfigureExcelRecordMap<TRow>(ExcelRecordMapBuilder<MapBuilderTestModel, TRow> builder)
                where TRow : IRowWriter
            {
                builder
                    .Column("Name", static (row, m) => row.Write(m.Name))
                    .Column("Age", static (row, m) => row.Write(m.Age))
                    .Column("BirthDate", static (row, m) => row.Write(m.BirthDate))
                    .Column("Active", static (row, m) => row.Write(m.Active))
                    .Column("Kind", static (row, m) => row.Write(m.Kind.ToString()));
            }
        }

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

            await using XlsxReader mappedReader = await Excel.FromXlsxAsync(ms, leaveOpen: true, ct: TestContext.Current.CancellationToken);
            List<MapBuilderTestModel> mapped = ExcelParser.Generated<MapBuilderTestModel>().Parse(mappedReader).ToList();

            ms.Position = 0;
            await using XlsxReader reflectionReader = await Excel.FromXlsxAsync(ms, ct: TestContext.Current.CancellationToken);
            List<ReflectionTestModel> reflected = ExcelParser.FromAttributes<ReflectionTestModel>().Parse(reflectionReader).ToList();

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

            await using XlsxReader mappedReader = await Excel.FromXlsxAsync(mappedStream, ct: TestContext.Current.CancellationToken);
            var mappedEnumerable = new ExcelEnumerable<MapBuilderTestModel>(mappedReader, new ExcelParserConfig(), info);
            ExcelParseException mappedEx = Assert.Throws<ExcelParseException>(() => mappedEnumerable.ToList());

            await using MemoryStream reflectedStream = await TypedWorkbook.BuildAsync(["Name"], ["Alice"]);
            await using XlsxReader reflectedReader = await Excel.FromXlsxAsync(reflectedStream, ct: TestContext.Current.CancellationToken);
            ExcelParseException reflectedEx = Assert.Throws<ExcelParseException>(
                () => ExcelParser.FromAttributes<RequiredAgeRow>().Parse(reflectedReader).ToList());

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
            await using (XlsxWorkbookWriter wb = XlsxWorkbookWriter.Create(mappedStream, leaveOpen: true))
            {
                XlsxSheetWriter sheet = wb.AddSheet("S1");
                await sheet.StartAsync(ct);
                var builder = new ExcelRecordMapBuilder<MapBuilderTestModel, XlsxRowWriter>();
                MapBuilderTestModel.ConfigureExcelRecordMap(builder);
                await WriteMappedRowAsync(sheet, builder, record, ct);
                await sheet.EndAsync(ct);
                await wb.EndAsync(ct);
            }

            await using var reflectedStream = new MemoryStream();
            await using (WorkbookRecordWriter<XlsxSheetWriter, XlsxRowWriter> writer =
                RecordWriter.CreateXlsx(reflectedStream, leaveOpen: true))
            {
                await writer.WriteSheetAsync("S1", [record], ct);
            }

            mappedStream.Position = 0;
            reflectedStream.Position = 0;
            await using XlsxReader mappedReader = await Excel.FromXlsxAsync(mappedStream, ct: TestContext.Current.CancellationToken);
            await using XlsxReader reflectedReader = await Excel.FromXlsxAsync(reflectedStream, ct: TestContext.Current.CancellationToken);
            List<ReflectionTestModel> mapped = ExcelParser.FromAttributes<ReflectionTestModel>().Parse(mappedReader).ToList();
            List<ReflectionTestModel> reflected = ExcelParser.FromAttributes<ReflectionTestModel>().Parse(reflectedReader).ToList();

            Assert.Single(mapped);
            Assert.Single(reflected);
            Assert.Equal(reflected[0].Name, mapped[0].Name);
            Assert.Equal(reflected[0].Age, mapped[0].Age);
            Assert.Equal(reflected[0].BirthDate, mapped[0].BirthDate);
            Assert.Equal(reflected[0].Active, mapped[0].Active);
            Assert.Equal(reflected[0].Kind, mapped[0].Kind);
        }

        private static async ValueTask WriteMappedRowAsync(XlsxSheetWriter sheet, ExcelRecordMapBuilder<MapBuilderTestModel, XlsxRowWriter> builder, MapBuilderTestModel record, CancellationToken ct)
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

        [Fact]
        public async Task MappedRecordWriterRoundTripsThroughXlsx()
        {
            CancellationToken ct = TestContext.Current.CancellationToken;
            var record = new MapBuilderTestModel { Name = "Carol", Age = 21, BirthDate = SampleBirthDate, Active = true, Kind = MapBuilderKind.Beta };

            await using var stream = new MemoryStream();
            await using (var writer = MappedRecordWriter.CreateXlsx(stream, leaveOpen: true))
            {
                await writer.WriteSheetAsync("S1", [record], ct);
            }

            stream.Position = 0;
            await using XlsxReader reader = await Excel.FromXlsxAsync(stream, ct: ct);
            List<MapBuilderTestModel> results = ExcelParser.Generated<MapBuilderTestModel>().Parse(reader).ToList();

            Assert.Single(results);
            AssertMatches(record, results[0]);
        }

        [Fact]
        public async Task MappedRecordWriterRoundTripsThroughXlsb()
        {
            CancellationToken ct = TestContext.Current.CancellationToken;
            var record = new MapBuilderTestModel { Name = "Dave", Age = 55, BirthDate = SampleBirthDate, Active = false, Kind = MapBuilderKind.Alpha };

            await using var stream = new MemoryStream();
            await using (var writer = MappedRecordWriter.CreateXlsb(stream, leaveOpen: true))
            {
                await writer.WriteSheetAsync("S1", [record], ct);
            }

            stream.Position = 0;
            await using XlsbReader reader = await Excel.FromXlsbAsync(stream, leaveOpen: false, ct: ct);
            List<MapBuilderTestModel> results = ExcelParser.Generated<MapBuilderTestModel>().Parse(reader).ToList();

            Assert.Single(results);
            AssertMatches(record, results[0]);
        }

        [Fact]
        public async Task MappedRecordWriterRoundTripsThroughXls()
        {
            CancellationToken ct = TestContext.Current.CancellationToken;
            var record = new MapBuilderTestModel { Name = "Erin", Age = 33, BirthDate = SampleBirthDate, Active = true, Kind = MapBuilderKind.Alpha };

            await using var stream = new MemoryStream();
            await using (var writer = MappedRecordWriter.CreateXls(stream, leaveOpen: true))
            {
                await writer.WriteSheetAsync("S1", [record], ct);
            }

            stream.Position = 0;
            await using XlsReader reader = Excel.FromXls(stream, leaveOpen: false);
            List<MapBuilderTestModel> results = ExcelParser.Generated<MapBuilderTestModel>().Parse(reader).ToList();

            Assert.Single(results);
            AssertMatches(record, results[0]);
        }

        [Fact]
        public async Task MappedRecordWriterRoundTripsThroughCsv()
        {
            CancellationToken ct = TestContext.Current.CancellationToken;
            var record = new MapBuilderTestModel { Name = "Frank", Age = 19, BirthDate = SampleBirthDate, Active = false, Kind = MapBuilderKind.Beta };

            await using var stream = new MemoryStream();
            await using (var writer = MappedRecordWriter.CreateCsv(stream, leaveOpen: true))
            {
                await writer.WriteSheetAsync("S1", [record], ct);
            }

            stream.Position = 0;
            await using CsvReader reader = Excel.FromCsv(stream, leaveOpen: false);
            List<MapBuilderTestModel> results = ExcelParser.Generated<MapBuilderTestModel>().Parse(reader).ToList();

            Assert.Single(results);
            Assert.Equal(record.Name, results[0].Name);
            Assert.Equal(record.Age, results[0].Age);
            Assert.Equal(record.Active, results[0].Active);
            Assert.Equal(record.Kind, results[0].Kind);
        }

        private sealed class ConfigureCountingModel : IExcelRecordMap<ConfigureCountingModel>
        {
            private static int _configurations;

            public string Name { get; set; } = "";

            internal static int Configurations => Volatile.Read(ref _configurations);

            public static void ConfigureExcelRecordMap<TRow>(ExcelRecordMapBuilder<ConfigureCountingModel, TRow> builder)
                where TRow : IRowWriter
            {
                Interlocked.Increment(ref _configurations);
                builder.Column("Name", static (row, m) => row.Write(m.Name));
            }
        }

        [Fact]
        public async Task RecordMapIsConfiguredOncePerRowWriterType()
        {
            CancellationToken ct = TestContext.Current.CancellationToken;
            var record = new ConfigureCountingModel { Name = "Grace" };

            await using var xlsxStream = new MemoryStream();
            await using (var xlsx = MappedRecordWriter.CreateXlsx(xlsxStream, leaveOpen: true))
            {
                await xlsx.WriteSheetAsync("S1", [record], ct);
            }
            Assert.Equal(1, ConfigureCountingModel.Configurations);

            await using var secondXlsxStream = new MemoryStream();
            await using (var again = MappedRecordWriter.CreateXlsx(secondXlsxStream, leaveOpen: true))
            {
                await again.WriteSheetAsync("S1", [record], ct);
            }
            Assert.Equal(1, ConfigureCountingModel.Configurations);

            await using var csvStream = new MemoryStream();
            await using (var csv = MappedRecordWriter.CreateCsv(csvStream, leaveOpen: true))
            {
                await csv.WriteSheetAsync("S1", [record], ct);
            }
            Assert.Equal(2, ConfigureCountingModel.Configurations);
        }

        private static void AssertMatches(MapBuilderTestModel expected, MapBuilderTestModel actual)
        {
            Assert.Equal(expected.Name, actual.Name);
            Assert.Equal(expected.Age, actual.Age);
            Assert.Equal(expected.BirthDate, actual.BirthDate);
            Assert.Equal(expected.Active, actual.Active);
            Assert.Equal(expected.Kind, actual.Kind);
        }
    }
}
