using System.Globalization;
using System.IO.Compression;
using System.Text;
using System.Xml.Linq;
using ExcelReader.Core.Reader;
using ExcelReader.Core.Reader.Xlsb;
using ExcelReader.Core.Writer;
using ExcelReader.Core.Writer.Csv;
using ExcelReader.Core.Writer.Xls;
using ExcelReader.Core.Writer.Xlsb;
using ExcelReader.Core.Writer.Xlsx;

namespace ExcelReader.Tests.Writer.Styles
{
    public class CellStyleWriteTests
    {
        private const int XlsBlankRecord = 0x0201;

        public static TheoryData<string> Formats => ["xlsx", "xlsb", "xls"];

        private static readonly CellStyle Bold = new() { Bold = true };
        private static readonly CellStyle Italic = new() { Italic = true };
        private static readonly CellStyle Fill = new() { FillColor = ExcelColor.FromRgb(0xFFFF00) };

        [Theory]
        [MemberData(nameof(Formats))]
        public void StyleAppliesToTheNextCellOnly(string format)
        {
            byte[] file = Build(format, book =>
            {
                int bold = book.AddStyle(Bold);
                using IRowWriter row = book.StartRow();
                row.Style(bold).Write(1);
                row.Write(2);
            });
            Assert.True(Probe(format, file, 0, 0).Bold);
            Assert.False(Probe(format, file, 0, 1).Bold);
        }

        [Theory]
        [MemberData(nameof(Formats))]
        public void CellStyleReplacesRowAndColumnStyle(string format)
        {
            byte[] file = Build(format, book =>
            {
                int italic = book.AddStyle(Italic);
                int fill = book.AddStyle(Fill);
                int bold = book.AddStyle(Bold);
                book.SetColumnStyle(0, italic);
                using IRowWriter row = book.StartStyledRow(fill);
                row.Style(bold).Write(1);
            });
            ProbedStyle style = Probe(format, file, 0, 0);
            Assert.True(style.Bold);
            Assert.False(style.Italic);
            Assert.Null(style.FillRgb);
        }

        [Theory]
        [MemberData(nameof(Formats))]
        public void StyleZeroForcesTheGeneralStyle(string format)
        {
            byte[] file = Build(format, book =>
            {
                int italic = book.AddStyle(Italic);
                book.SetColumnStyle(0, italic);
                using IRowWriter row = book.StartRow();
                row.Style(0).Write(1);
            });
            Assert.False(Probe(format, file, 0, 0).Italic);
        }

        [Theory]
        [MemberData(nameof(Formats))]
        public void NullWriteWithPendingStyleWritesAStyledBlank(string format)
        {
            byte[] file = Build(format, book =>
            {
                int fill = book.AddStyle(Fill);
                using IRowWriter row = book.StartRow();
                row.Style(fill).Write((int?)null);
            });
            Assert.Equal(0xFFFF00, ProbeBlank(format, file).FillRgb);
        }

        [Theory]
        [MemberData(nameof(Formats))]
        public void NullStringWithPendingStyleWritesAStyledBlank(string format)
        {
            byte[] file = Build(format, book =>
            {
                int fill = book.AddStyle(Fill);
                using IRowWriter row = book.StartRow();
                row.Style(fill).Write((string?)null);
                row.Write(1);
            });
            Assert.Equal(0xFFFF00, ProbeBlank(format, file).FillRgb);
            Assert.Null(Probe(format, file, 0, 1).FillRgb);
        }

        [Theory]
        [MemberData(nameof(Formats))]
        public void SkipWithPendingStyleThrows(string format)
        {
            Build(format, book =>
            {
                int bold = book.AddStyle(Bold);
                using IRowWriter row = book.StartRow();
                row.Style(bold);
                Assert.Throws<InvalidOperationException>(() => row.Skip(1));
                row.Write(1);
            });
        }

        [Theory]
        [MemberData(nameof(Formats))]
        public void SkipBeforeStyleIsAllowed(string format)
        {
            byte[] file = Build(format, book =>
            {
                int bold = book.AddStyle(Bold);
                using IRowWriter row = book.StartRow();
                row.Skip(1);
                row.Style(bold).Write(1);
            });
            Assert.True(Probe(format, file, 0, 1).Bold);
        }

        [Theory]
        [MemberData(nameof(Formats))]
        public void DisposingRowWithPendingStyleThrows(string format)
        {
            byte[] file = Build(format, book =>
            {
                int bold = book.AddStyle(Bold);
                IRowWriter row = book.StartRow();
                row.Write(1);
                row.Style(bold);
                Assert.Throws<InvalidOperationException>(row.Dispose);
                using IRowWriter next = book.StartRow();
                next.Write(2);
            });
            Assert.False(Probe(format, file, 0, 0).Bold);
            Assert.False(Probe(format, file, 1, 0).Bold);
        }

        [Theory]
        [MemberData(nameof(Formats))]
        public void DisposingRowAsyncWithPendingStyleThrows(string format)
        {
            byte[] file = Build(format, book =>
            {
                int bold = book.AddStyle(Bold);
                IRowWriter row = book.StartRow();
                row.Write(1);
                row.Style(bold);
                Assert.Throws<InvalidOperationException>(() => row.DisposeAsync().AsTask().GetAwaiter().GetResult());
                using IRowWriter next = book.StartRow();
                next.Write(2);
            });
            Assert.False(Probe(format, file, 1, 0).Bold);
        }

        [Theory]
        [MemberData(nameof(Formats))]
        public void DateUnderPendingStyleWithoutFormatReadsAsDate(string format)
        {
            byte[] file = Build(format, book =>
            {
                int bold = book.AddStyle(Bold);
                using IRowWriter row = book.StartRow();
                row.Style(bold).Write(new DateTime(2026, 1, 2));
            });
            Assert.Equal(CellType.Date, CellAt(format, file, 0, 0).Type);
            Assert.True(Probe(format, file, 0, 0).Bold);
        }

        [Theory]
        [MemberData(nameof(Formats))]
        public void UnknownStyleIdThrows(string format)
        {
            Build(format, book =>
            {
                using IRowWriter row = book.StartRow();
                Assert.Throws<ArgumentOutOfRangeException>(() => row.Style(999));
                Assert.Throws<ArgumentOutOfRangeException>(() => row.Style(-1));
            });
        }

        [Fact]
        public void CsvAcceptsAndIgnoresStyle()
        {
            using var ms = new MemoryStream();
            using (CsvWorkbookWriter workbook = CsvWorkbookWriter.Create(ms, leaveOpen: true))
            {
                CsvSheetWriter sheet = workbook.AddSheet("S1");
                using (CsvRowWriter row = sheet.StartRow())
                {
                    row.Style(0).Write("a");
                    row.Write("b");
                }
                sheet.End();
                workbook.End();
            }
            Assert.StartsWith("a,b", Encoding.UTF8.GetString(ms.ToArray()), StringComparison.Ordinal);
        }

        [Fact]
        public void CsvStyleRulesMatchTheOtherFormats()
        {
            using var ms = new MemoryStream();
            using CsvWorkbookWriter workbook = CsvWorkbookWriter.Create(ms, leaveOpen: true);
            CsvSheetWriter sheet = workbook.AddSheet("S1");
            CsvRowWriter row = sheet.StartRow();
            Assert.Throws<ArgumentOutOfRangeException>(() => row.Style(-1));
            row.Style(0);
            Assert.Throws<InvalidOperationException>(() => row.Skip(1));
            row.Write((int?)null);
            row.Style(0);
            Assert.Throws<InvalidOperationException>(row.Dispose);
        }

        [Fact]
        public void ExternalRowWriterDefaultStyleThrows()
        {
            IRowWriter row = new ExternalRowWriter();
            Assert.Throws<NotSupportedException>(() => row.Style(0));
        }

        private static byte[] Build(string format, Action<Book> script)
        {
            var ms = new MemoryStream();
            Book book = Open(format, ms);
            script(book);
            book.Finish();
            return ms.ToArray();
        }

        private static Book Open(string format, MemoryStream ms)
        {
            if (string.Equals(format, "xlsx", StringComparison.Ordinal))
            {
                return Book.Of<XlsxWorkbookWriter, XlsxSheetWriter, XlsxRowWriter>(XlsxWorkbookWriter.Create(ms, leaveOpen: true));
            }
            if (string.Equals(format, "xlsb", StringComparison.Ordinal))
            {
                return Book.Of<XlsbWorkbookWriter, XlsbSheetWriter, XlsbRowWriter>(XlsbWorkbookWriter.Create(ms, leaveOpen: true));
            }
            return Book.Of<XlsWorkbookWriter, XlsSheetWriter, XlsRowWriter>(XlsWorkbookWriter.Create(ms, leaveOpen: true));
        }

        private static ProbedStyle Probe(string format, byte[] file, int row, int column)
        {
            if (string.Equals(format, "xlsx", StringComparison.Ordinal))
            {
                return StyleProbe.Xlsx(file, row, column);
            }
            if (string.Equals(format, "xlsb", StringComparison.Ordinal))
            {
                return StyleProbe.Xlsb(file, row, column);
            }
            return StyleProbe.Xls(file, row, column);
        }

        // The readers skip blank cells, so a styled blank is found in the raw records.
        private static ProbedStyle ProbeBlank(string format, byte[] file)
        {
            if (string.Equals(format, "xlsx", StringComparison.Ordinal))
            {
                return StyleProbe.XlsxAt(file, XlsxBlankStyle(file));
            }
            if (string.Equals(format, "xlsb", StringComparison.Ordinal))
            {
                byte[] blank = StyleProbe.XlsbRecords(file, "xl/worksheets/sheet1.bin").First(r => r.Id == Brt.CellBlank).Payload;
                return StyleProbe.XlsbAt(file, blank[4] | (blank[5] << 8) | (blank[6] << 16));
            }
            byte[] xlsBlank = StyleProbe.XlsPayloads(StyleProbe.XlsRecords(file), XlsBlankRecord).First();
            return StyleProbe.XlsAt(file, BitConverter.ToUInt16(xlsBlank, 4));
        }

        private static int XlsxBlankStyle(byte[] file)
        {
            XNamespace main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
            using var zip = new ZipArchive(new MemoryStream(file), ZipArchiveMode.Read);
            XElement sheet = XDocument.Load(zip.GetEntry("xl/worksheets/sheet1.xml")!.Open()).Root!;
            XElement blank = sheet.Descendants(main + "c").First(c => !c.HasElements);
            return int.Parse(blank.Attribute("s")!.Value, CultureInfo.InvariantCulture);
        }

        private static Cell CellAt(string format, byte[] file, int row, int column)
        {
            using IExcelWorkbook reader = OpenReader(format, file);
            using var rows = reader.FirstSheet.GetEnumerator();
            for (int i = 0; i <= row; i++)
            {
                Assert.True(rows.MoveNext());
            }
            return rows.Current[column];
        }

        private static IExcelWorkbook OpenReader(string format, byte[] file)
        {
            var ms = new MemoryStream(file);
            if (string.Equals(format, "xlsx", StringComparison.Ordinal))
            {
                return Excel.FromXlsx(ms);
            }
            if (string.Equals(format, "xlsb", StringComparison.Ordinal))
            {
                return Excel.FromXlsb(ms);
            }
            return Excel.FromXls(ms);
        }

        private sealed class Book
        {
            internal Func<CellStyle, int> AddStyle = static _ => 0;
            internal Func<IRowWriter> StartRow = static () => throw new InvalidOperationException();
            internal Func<int, IRowWriter> StartStyledRow = static _ => throw new InvalidOperationException();
            internal Action<int, int> SetColumnStyle = static (_, _) => { };
            internal Action Finish = static () => { };

            internal static Book Of<TWorkbook, TSheet, TRow>(TWorkbook workbook)
                where TWorkbook : IWorkbookWriter<TSheet>
                where TSheet : ISheetWriter<TRow>
                where TRow : IRowWriter
            {
                TSheet sheet = workbook.AddSheet("S1");
                return new Book
                {
                    AddStyle = workbook.AddStyle,
                    StartRow = () => sheet.StartRow(),
                    StartStyledRow = styleId => sheet.StartRow(styleId),
                    SetColumnStyle = sheet.SetColumnStyle,
                    Finish = () =>
                    {
                        sheet.End();
                        workbook.End();
                        workbook.Dispose();
                    },
                };
            }
        }

        private sealed class ExternalRowWriter : IRowWriter
        {
            public void Write(string? value) { }
            public void Write(bool value) { }
            public void Write(bool? value) { }
            public void Write(DateTime value) { }
            public void Write(DateTime? value) { }
            public void Write(DateOnly value) { }
            public void Write(DateOnly? value) { }
            public void Write(TimeOnly value) { }
            public void Write(TimeOnly? value) { }
            public void Write<T>(T value) where T : IUtf8SpanFormattable { }
            public void Write<T>(T? value) where T : struct, IUtf8SpanFormattable { }
            public void Skip(int count = 1) { }
            public void Dispose() { }
            public ValueTask DisposeAsync()
            {
                return ValueTask.CompletedTask;
            }
        }
    }
}
