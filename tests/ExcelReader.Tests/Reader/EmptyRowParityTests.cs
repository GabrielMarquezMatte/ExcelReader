using ExcelReader.Core.Reader;
using ExcelReader.Core.Writer;
using ExcelReader.Core.Writer.Xls;
using ExcelReader.Core.Writer.Xlsb;
using ExcelReader.Core.Writer.Xlsx;
using ExcelReader.Tests.Reader.Xls;

namespace ExcelReader.Tests.Reader
{
    public sealed class EmptyRowParityTests
    {
        private const int RowRecord = 0x0208;
        private const int LabelRecord = 0x0204;

        private static void WriteSparseSheet<TSheet, TRow>(IWorkbookWriter<TSheet> workbook)
            where TSheet : ISheetWriter<TRow>
            where TRow : IRowWriter
        {
            TSheet sheet = workbook.AddSheet("s");
            string?[] values = ["a", null, "b", null];
            foreach (string? value in values)
            {
                using TRow row = sheet.StartRow();
                row.Write(value);
                row.Write((long?)null);
            }
            sheet.End();
            sheet.Dispose();
            workbook.End();
        }

        private static List<string?> FirstCells<TEnumerator>(TEnumerator e)
            where TEnumerator : IExcelRowEnumerator
        {
            using TEnumerator owned = e;
            var cells = new List<string?>();
            while (owned.MoveNext())
            {
                cells.Add(owned.Current.ColumnCount == 0 ? null : owned.Current[0].GetString());
            }
            return cells;
        }

        [Fact]
        public void Xlsx_Yields_Rows_Written_Without_Cells()
        {
            using var ms = new MemoryStream();
            using (XlsxWorkbookWriter workbook = XlsxWorkbookWriter.Create(ms, leaveOpen: true))
            {
                WriteSparseSheet<XlsxSheetWriter, XlsxRowWriter>(workbook);
            }

            using var reader = Excel.FromXlsx(ms.ToArray());
            Assert.Equal(["a", null, "b", null], FirstCells(reader.GetEnumerator()));
        }

        [Fact]
        public void Xlsb_Yields_Rows_Written_Without_Cells()
        {
            using var ms = new MemoryStream();
            using (XlsbWorkbookWriter workbook = XlsbWorkbookWriter.Create(ms, leaveOpen: true))
            {
                WriteSparseSheet<XlsbSheetWriter, XlsbRowWriter>(workbook);
            }

            using var reader = Excel.FromXlsb(ms.ToArray());
            Assert.Equal(["a", null, "b", null], FirstCells(reader.GetEnumerator()));
        }

        [Fact]
        public void Xls_Yields_Rows_Written_Without_Cells()
        {
            using var ms = new MemoryStream();
            using (XlsWorkbookWriter workbook = XlsWorkbookWriter.Create(ms, leaveOpen: true))
            {
                WriteSparseSheet<XlsSheetWriter, XlsRowWriter>(workbook);
            }

            using var reader = Excel.FromXls(ms.ToArray());
            Assert.Equal(["a", null, "b", null], FirstCells(reader.GetEnumerator()));
        }

        [Fact]
        public void Xls_Yields_A_Row_Record_Without_Cells_In_Row_Order()
        {
            using var ms = XlsWorkbookBuilder.BuildRawSheet(
                includeEof: true,
                (RowRecord, XlsWorkbookBuilder.RawRowOnly(0)),
                (RowRecord, XlsWorkbookBuilder.RawRowOnly(1)),
                (RowRecord, XlsWorkbookBuilder.RawRowOnly(2)),
                (RowRecord, XlsWorkbookBuilder.RawRowOnly(3)),
                (LabelRecord, XlsWorkbookBuilder.RawLabel(0, 0, "a")),
                (LabelRecord, XlsWorkbookBuilder.RawLabel(2, 0, "b")));

            using var reader = Excel.FromXls(ms);
            Assert.Equal(["a", null, "b", null], FirstCells(reader.GetEnumerator()));
        }

        [Fact]
        public void Xls_Does_Not_Repeat_A_Row_Whose_Row_Record_Follows_Its_Cells()
        {
            using var ms = XlsWorkbookBuilder.BuildRawSheet(
                includeEof: true,
                (LabelRecord, XlsWorkbookBuilder.RawLabel(0, 0, "a")),
                (LabelRecord, XlsWorkbookBuilder.RawLabel(1, 0, "b")),
                (RowRecord, XlsWorkbookBuilder.RawRowOnly(0)),
                (RowRecord, XlsWorkbookBuilder.RawRowOnly(1)));

            using var reader = Excel.FromXls(ms);
            Assert.Equal(["a", "b"], FirstCells(reader.GetEnumerator()));
        }
    }
}
