using System.Globalization;
using ExcelReader.Core.Parser;
using ExcelReader.Core.Parser.Internal;
using ExcelReader.Core.Reader;

namespace ExcelReader.Tests
{
    public class CsvHeaderBinderTests
    {
        private sealed class Person
        {
            public string? Name { get; set; }
            public int Age { get; set; }
        }

        [Fact]
        public void BindsTheHeaderAndReportsWhereDataStarts()
        {
            byte[] csv = "Name,Age\r\nAda,36\nGrace,45\n"u8.ToArray();
            using var reader = Excel.FromCsv(csv);

            CsvBoundColumnMap<Person> map = CsvHeaderBinder.Bind<Person>(
                reader,
                new ExcelParserConfig(),
                TypeMapper<Person>.GetCsvInfo(),
                out long firstDataRecordOffset);

            Assert.Equal(10, firstDataRecordOffset);
            Assert.Equal(2, map.FieldParsers.Length);
            Assert.NotNull(map.FieldParsers[0]);
            Assert.NotNull(map.FieldParsers[1]);
            Assert.Equal("Name", map.FieldNames[0]);
            Assert.Equal("Age", map.FieldNames[1]);
        }

        [Fact]
        public void AProjectorBuiltFromABoundMapParsesTheVeryFirstRowAsData()
        {
            byte[] all = "Name,Age\nAda,36\nGrace,45\n"u8.ToArray();
            using var headerReader = Excel.FromCsv(all);
            CsvBoundColumnMap<Person> map = CsvHeaderBinder.Bind<Person>(
                headerReader,
                new ExcelParserConfig(),
                TypeMapper<Person>.GetCsvInfo(),
                out long dataStart);

            // Feed the projector a source that begins at the first data record — no header in sight.
            using var dataReader = Excel.FromCsv(all.AsMemory((int)dataStart));
            using CsvReader.Enumerator rows = dataReader.GetEnumerator();
            var projector = new CsvRowProjector<Person>(
                TypeMapper<Person>.GetCsvInfo(),
                map,
                CultureInfo.InvariantCulture,
                throwOnParseFailure: false);

            var people = new List<Person>();
            while (rows.MoveNext())
            {
                Person model = null!;
                if (projector.Advance(rows, ref model) == ProjectionStep.Yield)
                {
                    people.Add(model);
                }
            }

            Assert.Equal(2, people.Count);
            Assert.Equal("Ada", people[0].Name);
            Assert.Equal(36, people[0].Age);
            Assert.Equal("Grace", people[1].Name);
            Assert.Equal(45, people[1].Age);
        }

        [Fact]
        public void TheSameBoundMapIsSafeToShareAcrossConcurrentProjectors()
        {
            byte[] all = "Name,Age\nAda,36\nGrace,45\n"u8.ToArray();
            using var headerReader = Excel.FromCsv(all);
            CsvBoundColumnMap<Person> map = CsvHeaderBinder.Bind<Person>(
                headerReader,
                new ExcelParserConfig(),
                TypeMapper<Person>.GetCsvInfo(),
                out long dataStart);

            var results = new List<Person>[8];
            Parallel.For(0, 8, i =>
            {
                using var dataReader = Excel.FromCsv(all.AsMemory((int)dataStart));
                using CsvReader.Enumerator rows = dataReader.GetEnumerator();
                var projector = new CsvRowProjector<Person>(
                    TypeMapper<Person>.GetCsvInfo(), map, CultureInfo.InvariantCulture, throwOnParseFailure: false);
                var local = new List<Person>();
                while (rows.MoveNext())
                {
                    Person model = null!;
                    if (projector.Advance(rows, ref model) == ProjectionStep.Yield)
                    {
                        local.Add(model);
                    }
                }
                results[i] = local;
            });

            Assert.All(results, r =>
            {
                Assert.Equal(2, r.Count);
                Assert.Equal("Ada", r[0].Name);
                Assert.Equal(45, r[1].Age);
            });
        }
    }
}
