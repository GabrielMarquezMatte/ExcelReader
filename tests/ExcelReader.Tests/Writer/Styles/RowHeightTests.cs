using System.Text;
using ExcelReader.Core.Writer;
using ExcelReader.Core.Writer.Csv;
using ExcelReader.Core.Writer.Xls;
using ExcelReader.Core.Writer.Xlsb;
using ExcelReader.Core.Writer.Xlsx;

namespace ExcelReader.Tests.Writer.Styles
{
    public class RowHeightTests
    {
        public static TheoryData<string> Formats => ["xlsx", "xlsb", "xls"];

        public static TheoryData<string, double> OutOfRange
        {
            get
            {
                var data = new TheoryData<string, double>();
                foreach (string format in new[] { "xlsx", "xlsb", "xls", "csv" })
                {
                    foreach (double height in new[] { 0, -1, 409.5, double.NaN, double.PositiveInfinity })
                    {
                        data.Add(format, height);
                    }
                }
                return data;
            }
        }

        [Theory]
        [MemberData(nameof(Formats))]
        public void RowHeightRoundTrips(string format)
        {
            byte[] file = Build(format, sheet =>
            {
                using (IRowWriter row = sheet.StartRow(0, 30.5))
                {
                    row.Write(1);
                }
                using IRowWriter next = sheet.StartRow(0);
                next.Write(2);
            });
            Assert.Equal(30.5, Height(format, file, 0));
            Assert.Null(Height(format, file, 1));
        }

        [Theory]
        [MemberData(nameof(Formats))]
        public void RowHeightRoundsToTheNearestTwip(string format)
        {
            byte[] file = Build(format, sheet =>
            {
                using IRowWriter row = sheet.StartRow(0, 12.34);
                row.Write(1);
            });
            Assert.Equal(12.35, Height(format, file, 0));
        }

        [Theory]
        [MemberData(nameof(Formats))]
        public void RowHeightWithStyleKeepsTheStyle(string format)
        {
            byte[] file = Build(format, sheet =>
            {
                int bold = sheet.AddStyle(new CellStyle { Bold = true });
                using IRowWriter row = sheet.StartRow(bold, 20);
                row.Write(1);
            });
            Assert.Equal(20, Height(format, file, 0));
            Assert.True(Bold(format, file));
        }

        [Theory]
        [MemberData(nameof(OutOfRange))]
        public async Task RowHeightOutOfRangeThrows(string format, double height)
        {
            Sheet sheet = Open(format, new MemoryStream());
            Assert.Throws<ArgumentOutOfRangeException>(() => sheet.StartRow(0, height));
            await Assert.ThrowsAsync<ArgumentOutOfRangeException>(async () => await sheet.StartRowAsync(0, height));
            sheet.Finish();
        }

        [Theory]
        [MemberData(nameof(Formats))]
        public void RowHeightUnknownStyleThrows(string format)
        {
            Build(format, sheet => Assert.Throws<ArgumentOutOfRangeException>(() => sheet.StartRow(999, 20)));
        }

        [Fact]
        public void CsvIgnoresRowHeight()
        {
            byte[] file = Build("csv", sheet =>
            {
                using IRowWriter row = sheet.StartRow(0, 30.5);
                row.Write("a");
            });
            Assert.StartsWith("a", Encoding.UTF8.GetString(file), StringComparison.Ordinal);
        }

        [Theory]
        [MemberData(nameof(Formats))]
        public async Task AsyncOverloadMatchesSync(string format)
        {
            var ms = new MemoryStream();
            Sheet sheet = Open(format, ms);
            int bold = sheet.AddStyle(new CellStyle { Bold = true });
            await using (IRowWriter first = await sheet.StartRowAsync(bold, 30.5))
            {
                first.Write(1);
            }
            await using (IRowWriter second = await sheet.StartRowAsync(0, 18))
            {
                second.Write(2);
            }
            sheet.Finish();
            byte[] file = ms.ToArray();
            Assert.Equal(30.5, Height(format, file, 0));
            Assert.Equal(18, Height(format, file, 1));
            Assert.True(Bold(format, file));
        }

        [Fact]
        public void XlsbRowSpanningColumnBlocksKeepsItsHeight()
        {
            byte[] file = Build("xlsb", sheet =>
            {
                using IRowWriter row = sheet.StartRow(0, 25);
                row.Write(1);
                row.Skip(1099);
                row.Write(2);
            });
            Assert.Equal(25, Height("xlsb", file, 0));
        }

        [Fact]
        public void XlsContinuationSheetKeepsTheHeight()
        {
            const int RowsPerSheet = 65536;
            byte[] file = Build("xls", sheet =>
            {
                for (int i = 0; i < RowsPerSheet; i++)
                {
                    using IRowWriter row = sheet.StartRow(0);
                    row.Write(i);
                }
                using IRowWriter tall = sheet.StartRow(0, 40);
                tall.Write(1);
            });
            double?[] heights = StyleProbe.XlsRowRecords(file).Select(StyleProbe.XlsRowHeight).Where(h => h is not null).ToArray();
            Assert.Equal([40.0], heights);
        }

        private static double? Height(string format, byte[] file, int row)
        {
            if (string.Equals(format, "xlsx", StringComparison.Ordinal))
            {
                return StyleProbe.RowHeightXlsx(file, row);
            }
            if (string.Equals(format, "xlsb", StringComparison.Ordinal))
            {
                return StyleProbe.RowHeightXlsb(file, row);
            }
            return StyleProbe.RowHeightXls(file, row);
        }

        private static bool Bold(string format, byte[] file)
        {
            if (string.Equals(format, "xlsx", StringComparison.Ordinal))
            {
                return StyleProbe.Xlsx(file, 0, 0).Bold;
            }
            if (string.Equals(format, "xlsb", StringComparison.Ordinal))
            {
                return StyleProbe.Xlsb(file, 0, 0).Bold;
            }
            return StyleProbe.Xls(file, 0, 0).Bold;
        }

        private static byte[] Build(string format, Action<Sheet> script)
        {
            var ms = new MemoryStream();
            Sheet sheet = Open(format, ms);
            script(sheet);
            sheet.Finish();
            return ms.ToArray();
        }

        private static Sheet Open(string format, MemoryStream ms)
        {
            if (string.Equals(format, "xlsx", StringComparison.Ordinal))
            {
                return Sheet.Of<XlsxWorkbookWriter, XlsxSheetWriter, XlsxRowWriter>(XlsxWorkbookWriter.Create(ms, leaveOpen: true));
            }
            if (string.Equals(format, "xlsb", StringComparison.Ordinal))
            {
                return Sheet.Of<XlsbWorkbookWriter, XlsbSheetWriter, XlsbRowWriter>(XlsbWorkbookWriter.Create(ms, leaveOpen: true));
            }
            if (string.Equals(format, "xls", StringComparison.Ordinal))
            {
                return Sheet.Of<XlsWorkbookWriter, XlsSheetWriter, XlsRowWriter>(XlsWorkbookWriter.Create(ms, leaveOpen: true));
            }
            return Sheet.Of<CsvWorkbookWriter, CsvSheetWriter, CsvRowWriter>(CsvWorkbookWriter.Create(ms, leaveOpen: true));
        }

        private sealed class Sheet
        {
            internal Func<CellStyle, int> AddStyle = static _ => 0;
            internal Func<int, IRowWriter> StartPlain = static _ => throw new InvalidOperationException();
            internal Func<int, double, IRowWriter> StartHeight = static (_, _) => throw new InvalidOperationException();
            internal Func<int, double, ValueTask<IRowWriter>> StartHeightAsync = static (_, _) => throw new InvalidOperationException();
            internal Action Finish = static () => { };

            internal IRowWriter StartRow(int styleId)
            {
                return StartPlain(styleId);
            }

            internal IRowWriter StartRow(int styleId, double height)
            {
                return StartHeight(styleId, height);
            }

            internal ValueTask<IRowWriter> StartRowAsync(int styleId, double height)
            {
                return StartHeightAsync(styleId, height);
            }

            internal static Sheet Of<TWorkbook, TSheet, TRow>(TWorkbook workbook)
                where TWorkbook : IWorkbookWriter<TSheet>
                where TSheet : ISheetWriter<TRow>
                where TRow : IRowWriter
            {
                TSheet sheet = workbook.AddSheet("S1");
                return new Sheet
                {
                    AddStyle = workbook.AddStyle,
                    StartPlain = styleId => sheet.StartRow(styleId),
                    StartHeight = (styleId, height) => sheet.StartRow(styleId, height),
                    StartHeightAsync = async (styleId, height) => await sheet.StartRowAsync(styleId, height),
                    Finish = () =>
                    {
                        sheet.End();
                        workbook.End();
                        workbook.Dispose();
                    },
                };
            }
        }
    }
}
