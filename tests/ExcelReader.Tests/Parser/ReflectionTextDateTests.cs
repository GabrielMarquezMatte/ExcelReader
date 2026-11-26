using ExcelReader.Core.Parser;
using ExcelReader.Core.Reader;

namespace ExcelReader.Tests.Parser
{
    public class ReflectionTextDateTests
    {
        private sealed class Row
        {
            public DateOnly Date { get; set; }

            public DateOnly? Optional { get; set; }

            public DateTime Moment { get; set; }

            public DateTime? OptionalMoment { get; set; }
        }

        [Fact]
        public async Task DatesStoredAsTextAreRead()
        {
            await using var ms = await TypedWorkbook.BuildAsync(
                ["Date", "Optional", "Moment", "OptionalMoment"],
                ["11/25/2026", "11/25/2026", "11/25/2026 13:45:00", "11/25/2026 13:45:00"]);
            await using var reader = await Excel.FromXlsxAsync(ms, ct: TestContext.Current.CancellationToken);

            Row row = ExcelParser.FromAttributes<Row>().Parse(reader.FirstSheet).Single();

            Assert.Equal(new DateOnly(2026, 11, 25), row.Date);
            Assert.Equal(new DateOnly(2026, 11, 25), row.Optional);
            Assert.Equal(new DateTime(2026, 11, 25, 13, 45, 0), row.Moment);
            Assert.Equal(new DateTime(2026, 11, 25, 13, 45, 0), row.OptionalMoment);
        }
    }
}
