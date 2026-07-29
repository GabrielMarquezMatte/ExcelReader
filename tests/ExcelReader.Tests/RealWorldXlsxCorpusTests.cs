using ExcelReader.Core.Enums;
using ExcelReader.Core.Reader;

namespace ExcelReader.Tests
{
    // Fixtures in this class are genuine binaries actually exported by the named producer — unlike
    // XlsxDialectShapeTests/XlsxProducerDialectShapeTests, which hand-author XML mimicking known
    // quirks. Keep fixtures tiny (a handful of rows) to bound repo size.
    public class RealWorldXlsxCorpusTests
    {
        // Generated via the SheetJS "xlsx" npm package (v0.18.5): XLSX.utils.aoa_to_sheet +
        // XLSX.writeFile({ cellDates: true }). Notably exercises two dialect quirks XlsxReader
        // supports specifically because non-Excel producers emit them: string cells typed t="str"
        // (normally a cached formula-result type) instead of t="s"/"inlineStr", and date cells typed
        // t="d" holding literal ISO-8601 text instead of a numeric serial.
        [Fact]
        public void ReadsSheetJsGeneratedWorkbook()
        {
            string path = Path.Combine(AppContext.BaseDirectory, "data", "sheetjs-sample.xlsx");
            using XlsxReader reader = Excel.FromFile(path);
            using XlsxReader.Enumerator e = reader.GetEnumerator();

            Assert.True(e.MoveNext());
            Assert.Equal("name", e.Current[0].GetString());
            Assert.Equal("quantity", e.Current[1].GetString());
            Assert.Equal("price", e.Current[2].GetString());
            Assert.Equal("in_stock", e.Current[3].GetString());
            Assert.Equal("restock_date", e.Current[4].GetString());
            // SheetJS writes plain strings as t="str" (normally a cached-formula-result type) rather
            // than t="s"/"inlineStr" — this is CellType.Formula here, not CellType.ExcelString.
            Assert.Equal(CellType.Formula, e.Current[0].Type);

            Assert.True(e.MoveNext());
            Assert.Equal("Widget", e.Current[0].GetString());
            Assert.True(e.Current[1].TryParse(null, out int quantity));
            Assert.Equal(12, quantity);
            Assert.True(e.Current[2].TryParse(null, out double price));
            Assert.Equal(4.5, price);
            Assert.Equal(CellType.Boolean, e.Current[3].Type);
            Assert.Equal("1", e.Current[3].GetString());
            Assert.Equal(CellType.Date, e.Current[4].Type);
            Assert.True(e.Current[4].TryGetDateTime(out DateTime restock));
            Assert.Equal(new DateTime(2024, 1, 14, 21, 0, 0, DateTimeKind.Unspecified), restock);

            Assert.True(e.MoveNext());
            Assert.Equal("Gadget", e.Current[0].GetString());
            Assert.True(e.Current[1].TryParse(null, out int gadgetQty));
            Assert.Equal(0, gadgetQty);
            Assert.True(e.Current[4].TryGetDateTime(out DateTime gadgetDate));
            Assert.Equal(new DateTime(2024, 2, 29, 21, 0, 0, DateTimeKind.Unspecified), gadgetDate);

            Assert.True(e.MoveNext());
            Assert.Equal("Gizmo", e.Current[0].GetString());
            Assert.True(e.Current[4].TryGetDateTime(out DateTime gizmoDate));
            Assert.Equal(new DateTime(2024, 6, 29, 21, 0, 0, DateTimeKind.Unspecified), gizmoDate);

            Assert.False(e.MoveNext());
        }
    }
}
