using System.Globalization;
using ExcelReader.Core;
using ExcelReader.Core.Reader;
using ExcelReader.Core.Reader.Csv;
using ExcelReader.Core.Reader.Xls;
using ExcelReader.Core.Reader.Xlsb;
using ExcelReader.Core.Reader.Xlsx;
using ExcelReader.Core.Writer;
using ExcelReader.Core.Writer.Xls;
using ExcelReader.Core.Writer.Xlsb;
using ExcelReader.Core.Writer.Xlsx;

namespace ExcelReader.Tests.Reader
{
    public sealed class WorkbookSheetTests
    {
        public enum Format
        {
            Xlsx,
            Xlsb,
            Xls,
        }

        private static readonly string[] _names = ["Alpha", "Beta", "Gamma"];
        private static readonly ExcelSheetVisibility[] _visibility =
            [ExcelSheetVisibility.Visible, ExcelSheetVisibility.Hidden, ExcelSheetVisibility.VeryHidden];

        private const int Rows = 50;

        private static void Fill<TSheet, TRow>(IWorkbookWriter<TSheet> workbook)
            where TSheet : ISheetWriter<TRow>
            where TRow : IRowWriter
        {
            for (int s = 0; s < _names.Length; s++)
            {
                TSheet sheet = workbook.AddSheet(_names[s], _visibility[s]);
                for (int r = 0; r < Rows; r++)
                {
                    using TRow row = sheet.StartRow();
                    row.Write(_names[s] + "-" + r.ToString(CultureInfo.InvariantCulture));
                    row.Write((s * 1000) + r);
                }
                sheet.End();
                sheet.Dispose();
            }
            workbook.End();
        }

        private static byte[] Build(Format format)
        {
            using MemoryStream buffer = new();
            switch (format)
            {
                case Format.Xlsx:
                    using (XlsxWorkbookWriter xlsx = XlsxWorkbookWriter.Create(buffer, leaveOpen: true))
                    {
                        Fill<XlsxSheetWriter, XlsxRowWriter>(xlsx);
                    }
                    break;
                case Format.Xlsb:
                    using (XlsbWorkbookWriter xlsb = XlsbWorkbookWriter.Create(buffer, leaveOpen: true))
                    {
                        Fill<XlsbSheetWriter, XlsbRowWriter>(xlsb);
                    }
                    break;
                default:
                    using (XlsWorkbookWriter xls = XlsWorkbookWriter.Create(buffer, leaveOpen: true))
                    {
                        Fill<XlsSheetWriter, XlsRowWriter>(xls);
                    }
                    break;
            }
            return buffer.ToArray();
        }

        private static IExcelWorkbook Open(Format format)
        {
            byte[] bytes = Build(format);
            return format switch
            {
                Format.Xlsx => Excel.FromXlsx(bytes),
                Format.Xlsb => Excel.FromXlsb(bytes),
                _ => Excel.FromXls(bytes),
            };
        }

        private static List<string> Drain(IExcelRowEnumerator e)
        {
            List<string> values = [];
            while (e.MoveNext())
            {
                values.Add(e.Current[0].GetString());
            }
            return values;
        }

        private static List<string> ExpectedRows(int sheet)
        {
            List<string> expected = [];
            for (int r = 0; r < Rows; r++)
            {
                expected.Add(_names[sheet] + "-" + r.ToString(CultureInfo.InvariantCulture));
            }
            return expected;
        }

        [Theory]
        [InlineData(Format.Xlsx)]
        [InlineData(Format.Xlsb)]
        [InlineData(Format.Xls)]
        public void Sheets_Carry_Index_Name_And_Visibility(Format format)
        {
            using IExcelWorkbook workbook = Open(format);
            Assert.Equal(3, workbook.SheetCount);
            for (int s = 0; s < 3; s++)
            {
                IExcelSheet sheet = workbook.SheetAt(s);
                Assert.Equal(s, sheet.Index);
                Assert.Equal(_names[s], sheet.Name);
                Assert.Equal(_visibility[s], sheet.Visibility);
                Assert.Equal(workbook.IsDate1904, sheet.IsDate1904);
            }
        }

        [Theory]
        [InlineData(Format.Xlsx)]
        [InlineData(Format.Xlsb)]
        [InlineData(Format.Xls)]
        public void Each_Sheet_Enumerates_Its_Own_Rows_In_Any_Order(Format format)
        {
            using IExcelWorkbook workbook = Open(format);
            foreach (int s in new[] { 2, 0, 1, 2 })
            {
                using IExcelRowEnumerator e = workbook.SheetAt(s).GetEnumerator();
                Assert.Equal(ExpectedRows(s), Drain(e), StringComparer.Ordinal);
            }
        }

        [Theory]
        [InlineData(Format.Xlsx)]
        [InlineData(Format.Xlsb)]
        [InlineData(Format.Xls)]
        public async Task Each_Sheet_Enumerates_Asynchronously(Format format)
        {
            using IExcelWorkbook workbook = Open(format);
            for (int s = 0; s < 3; s++)
            {
                List<string> values = [];
                await using IExcelRowEnumerator e = workbook.SheetAt(s).GetAsyncEnumerator(TestContext.Current.CancellationToken);
                while (await e.MoveNextAsync())
                {
                    values.Add(e.Current[0].GetString());
                }
                Assert.Equal(ExpectedRows(s), values, StringComparer.Ordinal);
            }
        }

        [Theory]
        [InlineData(Format.Xlsx)]
        [InlineData(Format.Xlsb)]
        [InlineData(Format.Xls)]
        public void Interleaved_Enumerators_Keep_Their_Own_Sheet(Format format)
        {
            using IExcelWorkbook workbook = Open(format);
            using IExcelRowEnumerator first = workbook.SheetAt(0).GetEnumerator();
            using IExcelRowEnumerator second = workbook.SheetAt(1).GetEnumerator();
            using IExcelRowEnumerator again = workbook.SheetAt(0).GetEnumerator();
            for (int r = 0; r < Rows; r++)
            {
                Assert.True(first.MoveNext());
                Assert.True(second.MoveNext());
                Assert.True(again.MoveNext());
                Assert.Equal(ExpectedRows(0)[r], first.Current[0].GetString());
                Assert.Equal(ExpectedRows(1)[r], second.Current[0].GetString());
                Assert.Equal(ExpectedRows(0)[r], again.Current[0].GetString());
            }
        }

        [Theory]
        [InlineData(Format.Xlsx)]
        [InlineData(Format.Xlsb)]
        [InlineData(Format.Xls)]
        public void TryGetSheet_Ignores_Case_And_Reports_A_Miss(Format format)
        {
            using IExcelWorkbook workbook = Open(format);
            Assert.True(workbook.TryGetSheet("bETA", out IExcelSheet? found));
            Assert.Equal(1, found.Index);
            Assert.Equal("Beta", found.Name);
            Assert.False(workbook.TryGetSheet("Delta", out IExcelSheet? missing));
            Assert.Null(missing);
        }

        [Theory]
        [InlineData(Format.Xlsx)]
        [InlineData(Format.Xlsb)]
        [InlineData(Format.Xls)]
        public void A_Sheet_Index_Out_Of_Range_Is_Rejected(Format format)
        {
            using IExcelWorkbook workbook = Open(format);
            Assert.Throws<ArgumentOutOfRangeException>(() => workbook.SheetAt(-1));
            Assert.Throws<ArgumentOutOfRangeException>(() => workbook.SheetAt(3));
        }

        [Theory]
        [InlineData(Format.Xlsx)]
        [InlineData(Format.Xlsb)]
        [InlineData(Format.Xls)]
        public void A_Disposed_Workbook_Refuses_Sheets_And_Enumerators(Format format)
        {
            IExcelWorkbook workbook = Open(format);
            IExcelSheet taken = workbook.SheetAt(0);
            workbook.Dispose();

            Assert.Throws<ObjectDisposedException>(() => workbook.SheetAt(0));
            Assert.Throws<ObjectDisposedException>(() => workbook.TryGetSheet("Alpha", out _));
            Assert.Throws<ObjectDisposedException>(() => taken.GetEnumerator());
            Assert.Throws<ObjectDisposedException>(() => taken.GetAsyncEnumerator(TestContext.Current.CancellationToken));
            Assert.Equal("Alpha", taken.Name);
        }

        [Fact]
        public void The_Typed_Xlsx_List_Indexes_Enumerates_And_Compares()
        {
            using XlsxWorkbook workbook = Excel.FromXlsx(Build(Format.Xlsx));
            ExcelSheetList<XlsxSheet> sheets = workbook.Sheets;

            Assert.Equal(3, sheets.Count);
            Assert.Same(sheets, workbook.Sheets);
            Assert.Equal(sheets[1], sheets[1]);
            Assert.NotEqual(sheets[0], sheets[1]);
            Assert.Throws<ArgumentOutOfRangeException>(() => sheets[3]);

            List<string> names = [];
            foreach (XlsxSheet sheet in sheets)
            {
                names.Add(sheet.Name);
            }
            Assert.Equal(_names, names, StringComparer.Ordinal);
            Assert.Equal(_names, sheets.Select(s => s.Name), StringComparer.Ordinal);

            Assert.True(workbook.TryGetSheet("gamma", out XlsxSheet gamma));
            using XlsxWorkbook.Enumerator e = gamma.GetEnumerator();
            Assert.Equal(ExpectedRows(2), Drain(e), StringComparer.Ordinal);

            workbook.Dispose();
            Assert.Throws<ObjectDisposedException>(() => workbook.Sheets);
        }

        [Fact]
        public void The_Typed_Xlsb_And_Xls_Lists_Hand_Out_Typed_Enumerators()
        {
            using XlsbWorkbook xlsb = Excel.FromXlsb(Build(Format.Xlsb));
            using XlsbWorkbook.Enumerator b = xlsb.Sheets[1].GetEnumerator();
            Assert.Equal(ExpectedRows(1), Drain(b), StringComparer.Ordinal);
            Assert.True(xlsb.TryGetSheet("ALPHA", out XlsbSheet alphaB));
            Assert.Equal(0, alphaB.Index);

            using XlsWorkbook xls = Excel.FromXls(Build(Format.Xls));
            using XlsWorkbook.Enumerator x = xls.Sheets[2].GetEnumerator();
            Assert.Equal(ExpectedRows(2), Drain(x), StringComparer.Ordinal);
            Assert.True(xls.TryGetSheet("beta", out XlsSheet betaX));
            Assert.Equal(ExcelSheetVisibility.Hidden, betaX.Visibility);
        }

        [Fact]
        public void A_Csv_Is_A_Workbook_With_One_Unnamed_Visible_Sheet()
        {
            using CsvReader csv = Excel.FromCsv("a,b\n1,2\n"u8.ToArray());
            IExcelWorkbook workbook = csv;

            Assert.Equal(1, workbook.SheetCount);
            Assert.Single(csv.Sheets);
            CsvSheet sheet = csv.Sheets[0];
            Assert.Equal(0, sheet.Index);
            Assert.Equal("", sheet.Name);
            Assert.Equal(ExcelSheetVisibility.Visible, sheet.Visibility);
            Assert.False(sheet.IsDate1904);
            Assert.True(workbook.TryGetSheet("", out IExcelSheet? unnamed));
            Assert.Equal(0, unnamed.Index);
            Assert.False(csv.TryGetSheet("Sheet1", out _));
            Assert.Throws<ArgumentOutOfRangeException>(() => workbook.SheetAt(1));

            using CsvReader.Enumerator e = sheet.GetEnumerator();
            Assert.True(e.MoveNext());
            Assert.Equal("a", e.Current[0].GetString());
            Assert.True(e.MoveNext());
            Assert.Equal("2", e.Current[1].GetString());
            Assert.False(e.MoveNext());
        }

        [Fact]
        public void A_Default_Sheet_Is_Inert_But_Does_Not_Crash_On_Inspection()
        {
            XlsxSheet sheet = default;
            Assert.Equal(0, sheet.Index);
            Assert.Null(sheet.Name);
            Assert.Equal(default, sheet);
            Assert.NotNull(sheet.ToString());
            Assert.Throws<InvalidOperationException>(() => sheet.GetEnumerator());
            Assert.Throws<InvalidOperationException>(() => default(XlsbSheet).GetEnumerator());
            Assert.Throws<InvalidOperationException>(() => default(XlsSheet).GetEnumerator());
            Assert.Throws<InvalidOperationException>(() => default(CsvSheet).GetEnumerator());
        }

        [Fact]
        public void FirstSheet_Is_The_Sheet_At_Index_Zero_On_Every_Workbook()
        {
            using XlsxWorkbook xlsx = Excel.FromXlsx(Build(Format.Xlsx));
            Assert.Equal(xlsx.Sheets[0], xlsx.FirstSheet);
            using XlsbWorkbook xlsb = Excel.FromXlsb(Build(Format.Xlsb));
            Assert.Equal(xlsb.Sheets[0], xlsb.FirstSheet);
            using XlsWorkbook xls = Excel.FromXls(Build(Format.Xls));
            Assert.Equal(xls.Sheets[0], xls.FirstSheet);
            using CsvReader csv = Excel.FromCsv("a\n"u8.ToArray());
            Assert.Equal(csv.Sheets[0], csv.FirstSheet);

            IExcelWorkbook agnostic = xlsx;
            Assert.Equal("Alpha", agnostic.FirstSheet.Name);
            using XlsxWorkbook.Enumerator e = xlsx.FirstSheet.GetEnumerator();
            Assert.Equal(ExpectedRows(0), Drain(e), StringComparer.Ordinal);

            xlsx.Dispose();
            Assert.Throws<ObjectDisposedException>(() => xlsx.FirstSheet);
            Assert.Throws<ObjectDisposedException>(() => agnostic.FirstSheet);
        }
    }
}
