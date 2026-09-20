using ExcelReader.Core.Enums;
using ExcelReader.Core.Reader;

namespace ExcelReader.Tests
{
    // XlsxDialectShapeTests/XlsxProducerDialectShapeTests, which hand-author XML mimicking known
    public class RealWorldXlsxCorpusTests
    {
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
