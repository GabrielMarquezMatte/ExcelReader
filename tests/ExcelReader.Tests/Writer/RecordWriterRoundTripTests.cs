using System.Globalization;
using ExcelReader.Core.Parser;
using ExcelReader.Core.Reader;
using ExcelReader.Core.Writer;
using ExcelReader.Core.Writer.Csv;
using ExcelReader.Core.Writer.Xls;
using ExcelReader.Core.Writer.Xlsb;
using ExcelReader.Core.Writer.Xlsx;

namespace ExcelReader.Tests.Writer
{
    public partial class RecordWriterRoundTripTests
    {
        public enum RoundTripKind { First, Second }

        [ExcelSerializable]
        public partial class TextFallbackModel
        {
            public Half Ratio { get; set; }
            public Half? RatioN { get; set; }
            public Int128 Big { get; set; }
            public UInt128 UBig { get; set; }
            public TimeSpan Duration { get; set; }
            public DateTimeOffset Stamp { get; set; }
            public DateTimeOffset? StampN { get; set; }
            public DateTimeOffset? Missing { get; set; }
            public Guid Id { get; set; }
            public char Letter { get; set; }
            public RoundTripKind Kind { get; set; }
        }

        public static TheoryData<ExcelFileFormat, bool> Cases()
        {
            var data = new TheoryData<ExcelFileFormat, bool>();
            foreach (ExcelFileFormat format in (ExcelFileFormat[])[ExcelFileFormat.Xlsx, ExcelFileFormat.Xlsb, ExcelFileFormat.Xls, ExcelFileFormat.Csv])
            {
                data.Add(format, false);
                data.Add(format, true);
            }
            return data;
        }

        [Theory]
        [MemberData(nameof(Cases))]
        public async Task EveryParsableTypeRoundTripsUnderACommaDecimalCulture(ExcelFileFormat format, bool generated)
        {
            CancellationToken ct = TestContext.Current.CancellationToken;
            var stamp = new DateTimeOffset(2026, 9, 23, 10, 30, 15, 123, TimeSpan.FromHours(-3));
            var expected = new TextFallbackModel
            {
                Ratio = (Half)1.5,
                RatioN = (Half)(-0.25),
                Big = Int128.MaxValue,
                UBig = UInt128.MaxValue,
                Duration = new TimeSpan(1, 2, 3, 4, 500),
                Stamp = stamp,
                StampN = stamp.AddDays(1),
                Missing = null,
                Id = Guid.Parse("11111111-2222-3333-4444-555555555555"),
                Letter = 'x',
                Kind = RoundTripKind.Second,
            };

            CultureInfo previous = CultureInfo.CurrentCulture;
            CultureInfo.CurrentCulture = CultureInfo.GetCultureInfo("pt-BR");
            using var ms = new MemoryStream();
            try
            {
                await WriteAsync(ms, format, generated, expected, ct);
            }
            finally
            {
                CultureInfo.CurrentCulture = previous;
            }

            ms.Position = 0;
            using IExcelRowReader reader = Excel.Open(ms, format);
            TextFallbackModel actual = generated
                ? ExcelParser.Generated<TextFallbackModel>().Parse(reader).Single()
                : ExcelParser.FromAttributes<TextFallbackModel>().Parse(reader).Single();

            Assert.Equal(expected.Ratio, actual.Ratio);
            Assert.Equal(expected.RatioN, actual.RatioN);
            Assert.Equal(expected.Big, actual.Big);
            Assert.Equal(expected.UBig, actual.UBig);
            Assert.Equal(expected.Duration, actual.Duration);
            Assert.Equal(expected.Stamp, actual.Stamp);
            Assert.Equal(expected.Stamp.Offset, actual.Stamp.Offset);
            Assert.Equal(expected.StampN, actual.StampN);
            Assert.Null(actual.Missing);
            Assert.Equal(expected.Id, actual.Id);
            Assert.Equal(expected.Letter, actual.Letter);
            Assert.Equal(expected.Kind, actual.Kind);
        }

        private static async Task WriteAsync(Stream stream, ExcelFileFormat format, bool generated, TextFallbackModel record, CancellationToken ct)
        {
            TextFallbackModel[] records = [record];
            ExcelRecordLayout<TextFallbackModel> layout = generated
                ? ExcelRecordLayout.Generated<TextFallbackModel>()
                : ExcelRecordLayout.FromAttributes<TextFallbackModel>();
            switch (format)
            {
                case ExcelFileFormat.Xlsx:
                    await using (var w = XlsxWorkbookWriter.Create(stream, leaveOpen: true)) await using (var s = w.AddSheet("S1")) { await s.WriteRecordsAsync(records, layout, ct); }
                    break;
                case ExcelFileFormat.Xlsb:
                    await using (var w = XlsbWorkbookWriter.Create(stream, leaveOpen: true)) await using (var s = w.AddSheet("S1")) { await s.WriteRecordsAsync(records, layout, ct); }
                    break;
                case ExcelFileFormat.Xls:
                    await using (var w = XlsWorkbookWriter.Create(stream, leaveOpen: true)) await using (var s = w.AddSheet("S1")) { await s.WriteRecordsAsync(records, layout, ct); }
                    break;
                default:
                    await using (var w = CsvWorkbookWriter.Create(stream, leaveOpen: true)) await using (var s = w.AddSheet("S1")) { await s.WriteRecordsAsync(records, layout, ct); }
                    break;
            }
        }
    }
}
