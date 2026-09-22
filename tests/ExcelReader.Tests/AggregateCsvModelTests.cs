using System.Globalization;
using System.Text;
using ExcelReader.Core.Parser;
using ExcelReader.Core.Parser.Internal;
using ExcelReader.Core.Reader;
using ExcelReader.Core.ValueObjects;
using ExcelReader.Core.Writer;

namespace ExcelReader.Tests
{
    public partial class AggregateCsvModelTests
    {
        [ExcelSerializable]
        public ref partial struct Person
        {
            [ExcelColumn("Name")]
            public ReadOnlySpan<byte> Name { get; set; }

            [ExcelColumn("Id")]
            [ExcelRequired]
            public int Id { get; set; }

            [ExcelColumn("Note")]
            public string? Note { get; set; }
        }

        public record struct IdOnly
        {
            public int Id { get; set; }
        }

        private sealed class PeopleLog : ICsvAccumulator<PeopleLog, Person>, ICsvAccumulator<PeopleLog, IdOnly>
        {
            public List<string> Rows { get; } = [];

            public void Add(Person model)
            {
                Rows.Add(Render(model));
            }

            public void Add(IdOnly model)
            {
                Rows.Add(model.Id.ToString(CultureInfo.InvariantCulture));
            }

            public void Merge(PeopleLog following)
            {
                Rows.AddRange(following.Rows);
            }
        }

        private static string Render(Person person)
        {
            return $"{Encoding.UTF8.GetString(person.Name)}|{person.Id}|{person.Note}";
        }

        private static byte[] PeopleCsv(int rows, int headerRow = 1, string header = "Name,Id,Note")
        {
            var sb = new StringBuilder();
            for (int i = 1; i < headerRow; i++)
            {
                sb.Append(CultureInfo.InvariantCulture, $"preamble {i}\n");
            }
            sb.Append(header).Append('\n');
            for (int i = 0; i < rows; i++)
            {
                string name = i % 5 == 0 ? $"\"multi\nline {i}, \"\"q\"\"\"" : $"name{i}";
                sb.Append(CultureInfo.InvariantCulture, $"{name},{i},note {i}\n");
            }
            return Encoding.UTF8.GetBytes(sb.ToString());
        }

        private static List<string> Sequential(byte[] csv, int headerRow)
        {
            using CsvReader reader = Excel.FromCsv(csv);
            var rows = new List<string>();
            foreach (Person person in new ExcelParser<Person>(new ExcelParserConfig { HeaderRow = headerRow }).Parse(reader))
            {
                rows.Add(Render(person));
            }
            return rows;
        }

        [System.Diagnostics.CodeAnalysis.UnconditionalSuppressMessage("Trimming", "IL2026", Justification = "Test")]
        [System.Diagnostics.CodeAnalysis.UnconditionalSuppressMessage("AOT", "IL3050", Justification = "Test")]
        private static CsvModelMap<Person>[] EveryMap()
        {
            return
            [
                CsvModelMap.Generated<Person>(),
                CsvModelMap.FromAttributes<Person>(),
                CsvModelMap.Build<Person>(static b => b
                    .PropertyRaw(["Name"], static (ref Person m, in Cell c, bool d, IFormatProvider pr) =>
                    {
                        m.Name = c.Value;
                        return true;
                    })
                    .Property(["Id"], ExcelCellReaders.Parsable, static (ref Person m, int v) => m.Id = v, isRequired: true, requireValue: true)
                    .Property(["Note"], ExcelCellReaders.String, static (ref Person m, string v) => m.Note = v)),
            ];
        }

        private static Task<PeopleLog> Chunked(byte[] csv, CsvModelMap<Person> map, int dop, int chunkSize, int headerRow)
        {
            return ParallelCsvProcessor.RunWithChunkSizeAsync(
                csv.AsMemory(),
                MappedAggregation<PeopleLog, Person>.Unbound,
                MappedAggregation<PeopleLog, Person>.Binder(map, headerRow),
                new CsvParallelOptions { DegreeOfParallelism = dop, HeaderRow = headerRow },
                chunkSize,
                TestContext.Current.CancellationToken);
        }

        private static Task<PeopleLog> Chunked(byte[] csv, CsvModelMap<IdOnly> map, int dop, int chunkSize, int headerRow)
        {
            return ParallelCsvProcessor.RunWithChunkSizeAsync(
                csv.AsMemory(),
                MappedAggregation<PeopleLog, IdOnly>.Unbound,
                MappedAggregation<PeopleLog, IdOnly>.Binder(map, headerRow),
                new CsvParallelOptions { DegreeOfParallelism = dop, HeaderRow = headerRow },
                chunkSize,
                TestContext.Current.CancellationToken);
        }

        private static byte[] IndexCsv(int rows, int headerRow)
        {
            var sb = new StringBuilder();
            if (headerRow == 1)
            {
                sb.Append("Name,Id\n");
            }
            for (int i = 0; i < rows; i++)
            {
                string name = i % 5 == 0 ? $"\"multi\nline {i}, \"\"q\"\"\"" : $"name{i}";
                sb.Append(CultureInfo.InvariantCulture, $"{name},{i}\n");
            }
            return Encoding.UTF8.GetBytes(sb.ToString());
        }

        [Fact]
        public async Task EveryMapSourceMatchesSequentialAcrossDegreeChunkSizeAndHeaderRow()
        {
            foreach (int headerRow in new[] { 1, 2 })
            {
                byte[] csv = PeopleCsv(rows: 40, headerRow);
                List<string> expected = Sequential(csv, headerRow);
                Assert.Equal(40, expected.Count);
                foreach (CsvModelMap<Person> map in EveryMap())
                {
                    foreach (int dop in new[] { 2, 3, 4, 8 })
                    {
                        foreach (int chunkSize in new[] { 1, 2, 3, 7, 16, 64 })
                        {
                            Assert.Equal(expected, (await Chunked(csv, map, dop, chunkSize, headerRow)).Rows);
                        }
                    }
                }
            }
        }

        [Fact]
        public async Task EveryPublicOverloadMatchesSequential()
        {
            byte[] csv = PeopleCsv(rows: 60_000);
            List<string> expected = Sequential(csv, headerRow: 1);
            var options = new CsvParallelOptions { DegreeOfParallelism = 4, HeaderRow = 1 };
            CancellationToken ct = TestContext.Current.CancellationToken;
            CsvModelMap<Person> map = CsvModelMap.Generated<Person>();
            string path = Path.Combine(Path.GetTempPath(), $"exr-model-{Guid.NewGuid():N}.csv");
            await File.WriteAllBytesAsync(path, csv, ct);
            try
            {
                Assert.Equal(expected, (await Excel.AggregateCsvParallelAsync<PeopleLog, Person>(csv.AsMemory(), map, options, ct)).Rows);
                Assert.Equal(expected, (await Excel.AggregateCsvParallelAsync<PeopleLog, Person>(path, map, options, ct)).Rows);
                await using (var file = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read, 4096, useAsync: true))
                {
                    Assert.Equal(expected, (await Excel.AggregateCsvParallelAsync<PeopleLog, Person>(file, map, options, ct)).Rows);
                }
                using var unseekable = new BufferedStream(new MemoryStream(csv, writable: false));
                Assert.Equal(expected, (await Excel.AggregateCsvParallelAsync<PeopleLog, Person>(unseekable, map, options, ct)).Rows);
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public async Task BindsReorderedAndAliasedColumns()
        {
            byte[] csv = "Note,Identifier,Name\nn1,1,ada\nn2,2,bob\n"u8.ToArray();
            CsvModelMap<Person> map = CsvModelMap.Build<Person>(static b => b
                .PropertyRaw(["Name"], static (ref Person m, in Cell c, bool d, IFormatProvider pr) =>
                {
                    m.Name = c.Value;
                    return true;
                })
                .Property(["Identifier", "Id"], ExcelCellReaders.Parsable, static (ref Person m, int v) => m.Id = v)
                .Property(["Note"], ExcelCellReaders.String, static (ref Person m, string v) => m.Note = v));

            PeopleLog log = await Excel.AggregateCsvParallelAsync<PeopleLog, Person>(csv, map, new CsvParallelOptions { HeaderRow = 1 }, TestContext.Current.CancellationToken);

            Assert.Equal(["ada|1|n1", "bob|2|n2"], log.Rows);
        }

        [Fact]
        public async Task AMissingRequiredColumnThrows()
        {
            byte[] csv = "Name,Note\nada,n1\n"u8.ToArray();
            await Assert.ThrowsAsync<ExcelParseException>(() => Excel.AggregateCsvParallelAsync<PeopleLog, Person>(
                csv, CsvModelMap.Generated<Person>(), new CsvParallelOptions { HeaderRow = 1 }, TestContext.Current.CancellationToken));
        }

        [Fact]
        public async Task ANameMapWithoutAHeaderRowIsRejectedAndAnIndexMapIsNot()
        {
            byte[] csv = "ada,1\nbob,2\n"u8.ToArray();
            CancellationToken ct = TestContext.Current.CancellationToken;
            await Assert.ThrowsAsync<ArgumentException>(() => Excel.AggregateCsvParallelAsync<PeopleLog, Person>(
                csv, CsvModelMap.Generated<Person>(), new CsvParallelOptions { HeaderRow = 0 }, ct));

            await Assert.ThrowsAsync<ArgumentNullException>(() => Excel.AggregateCsvParallelAsync<PeopleLog, Person>(
                csv, (CsvModelMap<Person>)null!, new CsvParallelOptions { HeaderRow = 1 }, ct));

            CsvModelMap<IdOnly> byIndex = CsvModelMap.Build<IdOnly>(static b => b.PropertyAt(1, ExcelCellReaders.Parsable, static (ref IdOnly m, int v) => m.Id = v));
            PeopleLog log = await Excel.AggregateCsvParallelAsync<PeopleLog, IdOnly>(csv, byIndex, new CsvParallelOptions { HeaderRow = 0 }, ct);
            Assert.Equal(["1", "2"], log.Rows);
        }

        [Fact]
        public async Task AnIndexMapRunsThroughThePartitionedPathForBothHeaderRowSettings()
        {
            CsvModelMap<IdOnly> byIndex = CsvModelMap.Build<IdOnly>(static b => b.PropertyAt(1, ExcelCellReaders.Parsable, static (ref IdOnly m, int v) => m.Id = v));
            foreach (int headerRow in new[] { 0, 1 })
            {
                byte[] csv = IndexCsv(rows: 200, headerRow);
                List<string> expected = [.. Enumerable.Range(0, 200).Select(static i => i.ToString(CultureInfo.InvariantCulture))];
                foreach (int chunkSize in new[] { 1, 7, 64 })
                {
                    Assert.Equal(expected, (await Chunked(csv, byIndex, dop: 4, chunkSize, headerRow)).Rows);
                }
            }
        }

        [Fact]
        public async Task SkipsEmptyRecords()
        {
            byte[] csv = "Name,Id,Note\nada,1,n\n\nbob,2,m\n"u8.ToArray();
            PeopleLog log = await Excel.AggregateCsvParallelAsync<PeopleLog, Person>(
                csv, CsvModelMap.Generated<Person>(), new CsvParallelOptions { HeaderRow = 1 }, TestContext.Current.CancellationToken);
            Assert.Equal(["ada|1|n", "bob|2|m"], log.Rows);
        }

        [Fact]
        public async Task ReportsTheExactRowSequentiallyAndZeroWhenPartitioned()
        {
            byte[] csv = "Name,Id,Note\nada,1,n\nbob,2,n\ncid,oops,n\ndan,4,n\n"u8.ToArray();
            CsvModelMap<Person> map = CsvModelMap.Generated<Person>(new ExcelParserConfig { ThrowOnParseFailure = true });

            ExcelParseException sequential = await Assert.ThrowsAsync<ExcelParseException>(() => Excel.AggregateCsvParallelAsync<PeopleLog, Person>(
                csv, map, new CsvParallelOptions { DegreeOfParallelism = 1, HeaderRow = 1 }, TestContext.Current.CancellationToken));
            Assert.Equal(4, sequential.Row);

            ExcelParseException partitioned = await Assert.ThrowsAsync<ExcelParseException>(() => Chunked(csv, map, dop: 4, chunkSize: 8, headerRow: 1));
            Assert.Equal(0, partitioned.Row);
        }

        [Fact]
        public async Task DiscardsParseFailuresFromAMisguessedPartition()
        {
            byte[] csv = "Name,Id,Note\nok,1,\"one\nBOOM,notint,x\"\nok,2,two\nok,3,\"three\nBOOM,bad,y\"\nok,4,four\n"u8.ToArray();
            CsvModelMap<Person> map = CsvModelMap.Generated<Person>(new ExcelParserConfig { ThrowOnParseFailure = true });
            List<string> expected = Sequential(csv, headerRow: 1);
            for (int chunkSize = 1; chunkSize <= csv.Length; chunkSize++)
            {
                Assert.Equal(expected, (await Chunked(csv, map, dop: 4, chunkSize, headerRow: 1)).Rows);
            }
        }

        [Fact]
        public async Task RequiredValueTrackingIsNotSharedAcrossWorkers()
        {
            byte[] csv = PeopleCsv(rows: 50_000);
            List<string> expected = Sequential(csv, headerRow: 1);
            Assert.Equal(expected, (await Chunked(csv, CsvModelMap.Generated<Person>(), dop: 8, chunkSize: 2048, headerRow: 1)).Rows);
        }

        [Fact]
        public async Task GeneratedRecordMapWritesTheSpanAsText()
        {
            CancellationToken ct = TestContext.Current.CancellationToken;
            await using var stream = new MemoryStream();
            await using (XlsxWorkbookWriter wb = await XlsxWorkbookWriter.CreateAsync(stream, leaveOpen: true, ct: ct))
            {
                await wb.StartAsync(ct);
                XlsxSheetWriter sheet = wb.AddSheet("S1");
                await sheet.StartAsync(ct);
                var builder = new ExcelRecordMapBuilder<Person, XlsxRowWriter>();
                Person.ConfigureExcelRecordMap(builder);
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
                    WritePerson(builder, row);
                }
                await sheet.EndAsync(ct);
                await wb.EndAsync(ct);
            }

            stream.Position = 0;
            await using XlsxReader reader = await Excel.FromXlsxAsync(stream, ct: ct);
            Assert.Equal(["Ada Lovelace|7|first"], ReadPeople(reader));
        }

        private static void WritePerson(ExcelRecordMapBuilder<Person, XlsxRowWriter> builder, XlsxRowWriter row)
        {
            builder.WriteRow(row, new Person { Name = "Ada Lovelace"u8, Id = 7, Note = "first" });
        }

        private static List<string> ReadPeople(XlsxReader reader)
        {
            var rows = new List<string>();
            foreach (Person person in new ExcelParser<Person>().Parse(reader))
            {
                rows.Add(Render(person));
            }
            return rows;
        }

        public struct WideRow : IEquatable<WideRow>
        {
            public int Last { get; set; }

            public override bool Equals(object? obj) => obj is WideRow row && Equals(row);
            public bool Equals(WideRow other) => Last == other.Last;
            public override int GetHashCode() => Last.GetHashCode();
            public static bool operator ==(WideRow left, WideRow right) => left.Equals(right);
            public static bool operator !=(WideRow left, WideRow right) => !left.Equals(right);
        }

        private sealed class WideLog : ICsvAccumulator<WideLog, WideRow>
        {
            public List<int> Values { get; } = [];

            public void Add(WideRow model)
            {
                Values.Add(model.Last);
            }

            public void Merge(WideLog following)
            {
                Values.AddRange(following.Values);
            }
        }

        [Fact]
        public async Task ATrackedMapWiderThanTheStackallocLimitStillBinds()
        {
            const int columns = 300;
            var header = new StringBuilder();
            var row = new StringBuilder();
            for (int index = 0; index < columns; index++)
            {
                header.Append(CultureInfo.InvariantCulture, $"c{index}");
                row.Append(index == columns - 1 ? "7" : "x");
                if (index < columns - 1)
                {
                    header.Append(',');
                    row.Append(',');
                }
            }
            byte[] csv = Encoding.UTF8.GetBytes($"{header}\n{row}\n{row}\n");

            CsvModelMap<WideRow> map = CsvModelMap.Build<WideRow>(b =>
            {
                for (int index = 0; index < columns - 1; index++)
                {
                    b.Property([$"c{index}"], ExcelCellReaders.String, static (ref WideRow m, string v) => { }, isRequired: true, requireValue: true);
                }
                b.Property([$"c{columns - 1}"], ExcelCellReaders.Parsable, static (ref WideRow m, int v) => m.Last = v, isRequired: true, requireValue: true);
            });

            WideLog log = await Excel.AggregateCsvParallelAsync<WideLog, WideRow>(
                csv, map, new CsvParallelOptions { HeaderRow = 1 }, TestContext.Current.CancellationToken);

            Assert.Equal([7, 7], log.Values);
        }
    }
}
