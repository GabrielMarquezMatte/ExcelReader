using ExcelReader.Core.Parser;
using ExcelReader.Core.Reader;
using ExcelReader.Core.Writer;
using ExcelReader.Core.Writer.Xls;
using ExcelReader.Core.Writer.Xlsb;
using ExcelReader.Core.Writer.Xlsx;

namespace ExcelReader.Tests.Parser
{
    /// <summary>
    /// A typed property binds the same way whichever format the cell came from. XLSX carries a number as
    /// text; XLSB and XLS carry it in binary, with nothing in <see cref="Cell.Value"/>.
    /// </summary>
    public sealed class TypedCellFormatParityTests
    {
        public static TheoryData<string> Formats => ["xlsx", "xlsb", "xls"];

        private sealed class Flags
        {
            public bool On { get; set; }
            public bool? Maybe { get; set; }
        }

        private sealed class Spans
        {
            public TimeSpan Days { get; set; }
            public TimeSpan? MaybeDays { get; set; }
        }

        private sealed class Shifts
        {
            public TimeOnly Start { get; set; }
            public TimeOnly? End { get; set; }
        }

        private sealed class Days
        {
            public DayOfWeek Day { get; set; }
        }

        [Theory]
        [MemberData(nameof(Formats))]
        public void ANumericZeroOrOneBindsToBool(string format)
        {
            byte[] bytes = Write(format, ["On", "Maybe"], row =>
            {
                row.Write(1);
                row.Write(0);
            }, row =>
            {
                row.Write(0.0);
                row.Write(1.0);
            }, row =>
            {
                row.Write(2);
                row.Write(0.5);
            });

            List<Flags> rows = Parse<Flags>(bytes);

            Assert.Equal([true, false, false], rows.Select(r => r.On));
            Assert.Equal([false, true, null], rows.Select(r => r.Maybe));
        }

        [Theory]
        [MemberData(nameof(Formats))]
        public void ANumberOtherThanZeroOrOneIsNotABool(string format)
        {
            byte[] bytes = Write(format, ["On", "Maybe"], row =>
            {
                row.Write(2);
                row.Write(1);
            });
            var strict = new ExcelParserOptions { ThrowOnParseFailure = true };

            using IExcelWorkbook workbook = Excel.Open(bytes);
            Assert.Throws<ExcelParseException>(() =>
            {
                foreach (Flags _ in ExcelParser.FromAttributes<Flags>(strict).Parse(workbook.FirstSheet))
                {
                }
            });
        }

        [Theory]
        [MemberData(nameof(Formats))]
        public void ANumberBindsToAParsableType(string format)
        {
            byte[] bytes = Write(format, ["Days", "MaybeDays"], row =>
            {
                row.Write(3);
                row.Write(300);
            });

            Spans row = Assert.Single(Parse<Spans>(bytes));

            Assert.Equal(TimeSpan.FromDays(3), row.Days);
            Assert.Equal(TimeSpan.FromDays(300), row.MaybeDays);
        }

        [Theory]
        [MemberData(nameof(Formats))]
        public void ATimeBindsFromItsFractionOrItsText(string format)
        {
            byte[] bytes = Write(format, ["Start", "End"], row =>
            {
                row.Write(new TimeOnly(8, 30));
                row.Write("17:45");
            }, row =>
            {
                row.Write("09:15");
                row.Write(0.75);
            });

            List<Shifts> rows = Parse<Shifts>(bytes);

            Assert.Equal([new TimeOnly(8, 30), new TimeOnly(9, 15)], rows.Select(r => r.Start));
            Assert.Equal([new TimeOnly(17, 45), new TimeOnly(18, 0)], rows.Select(r => r.End));
        }

        [Theory]
        [MemberData(nameof(Formats))]
        public void ADateStyledNumberBindsToAnEnumByValue(string format)
        {
            byte[] bytes = Write(format, ["Day"], row => row.Write(new DateTime(1900, 1, 1)), row => row.Write(5));

            Assert.Equal([DayOfWeek.Monday, DayOfWeek.Friday], Parse<Days>(bytes).Select(r => r.Day));
        }

        private static List<T> Parse<T>(byte[] bytes)
            where T : new()
        {
            using IExcelWorkbook workbook = Excel.Open(bytes);
            var rows = new List<T>();
            foreach (T row in ExcelParser.FromAttributes<T>().Parse(workbook.FirstSheet))
            {
                rows.Add(row);
            }
            return rows;
        }

        private static byte[] Write(string format, string[] header, params Action<IRowWriter>[] rows)
        {
            using MemoryStream ms = new();
            switch (format)
            {
                case "xlsx":
                    using (XlsxWorkbookWriter wb = XlsxWorkbookWriter.Create(ms, leaveOpen: true))
                    {
                        Fill<XlsxSheetWriter, XlsxRowWriter>(wb.AddSheet("S1"), header, rows);
                        wb.End();
                    }
                    break;
                case "xlsb":
                    using (XlsbWorkbookWriter wb = XlsbWorkbookWriter.Create(ms, leaveOpen: true))
                    {
                        Fill<XlsbSheetWriter, XlsbRowWriter>(wb.AddSheet("S1"), header, rows);
                        wb.End();
                    }
                    break;
                default:
                    using (XlsWorkbookWriter wb = XlsWorkbookWriter.Create(ms, leaveOpen: true))
                    {
                        Fill<XlsSheetWriter, XlsRowWriter>(wb.AddSheet("S1"), header, rows);
                        wb.End();
                    }
                    break;
            }
            return ms.ToArray();
        }

        private static void Fill<TSheet, TRow>(TSheet sheet, string[] header, Action<IRowWriter>[] rows)
            where TSheet : ISheetWriter<TRow>
            where TRow : IRowWriter, IDisposable
        {
            using (TRow row = sheet.StartRow())
            {
                foreach (string name in header)
                {
                    row.Write(name);
                }
            }
            foreach (Action<IRowWriter> write in rows)
            {
                using TRow row = sheet.StartRow();
                write(row);
            }
            sheet.End();
        }
    }
}
