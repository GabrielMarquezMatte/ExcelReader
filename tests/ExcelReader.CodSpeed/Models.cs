using ExcelReader.Core.Parser;

namespace ExcelReader.CodSpeed
{
    // Row shape written by the write scenarios: one text, one integer, one date, one float column.
    public sealed class WriteRecord
    {
        public string? Name { get; set; }
        public int Id { get; set; }
        public DateTime Date { get; set; }
        public double Value { get; set; }
    }

    // Typed projection of the 65K-row dataset's 14 columns. Header names carry spaces, so the
    // columns that do not match a property name bind through [ExcelColumn].
    public sealed class SalesRecord
    {
        public string? Region { get; set; }
        public string? Country { get; set; }

        [ExcelColumn("Item Type")]
        public string? ItemType { get; set; }

        [ExcelColumn("Sales Channel")]
        public string? SalesChannel { get; set; }

        [ExcelColumn("Order Priority")]
        public string? OrderPriority { get; set; }

        [ExcelColumn("Order Date")]
        public DateTime OrderDate { get; set; }

        [ExcelColumn("Order ID")]
        public long OrderId { get; set; }

        [ExcelColumn("Ship Date")]
        public DateTime ShipDate { get; set; }

        [ExcelColumn("Units Sold")]
        public int UnitsSold { get; set; }

        [ExcelColumn("Unit Price")]
        public double UnitPrice { get; set; }

        [ExcelColumn("Unit Cost")]
        public double UnitCost { get; set; }

        [ExcelColumn("Total Revenue")]
        public double TotalRevenue { get; set; }

        [ExcelColumn("Total Cost")]
        public double TotalCost { get; set; }

        [ExcelColumn("Total Profit")]
        public double TotalProfit { get; set; }
    }

    // Struct twin of SalesRecord. ExcelParser<T> binds columns through `ref TModel`, so parsing into a
    // struct skips the per-row model allocation a class T requires — the parse-xlsx-struct scenario
    // measures that path.
    public struct SalesRecordStruct
    {
        public string? Region { get; set; }
        public string? Country { get; set; }

        [ExcelColumn("Item Type")]
        public string? ItemType { get; set; }

        [ExcelColumn("Sales Channel")]
        public string? SalesChannel { get; set; }

        [ExcelColumn("Order Priority")]
        public string? OrderPriority { get; set; }

        [ExcelColumn("Order Date")]
        public DateTime OrderDate { get; set; }

        [ExcelColumn("Order ID")]
        public long OrderId { get; set; }

        [ExcelColumn("Ship Date")]
        public DateTime ShipDate { get; set; }

        [ExcelColumn("Units Sold")]
        public int UnitsSold { get; set; }

        [ExcelColumn("Unit Price")]
        public double UnitPrice { get; set; }

        [ExcelColumn("Unit Cost")]
        public double UnitCost { get; set; }

        [ExcelColumn("Total Revenue")]
        public double TotalRevenue { get; set; }

        [ExcelColumn("Total Cost")]
        public double TotalCost { get; set; }

        [ExcelColumn("Total Profit")]
        public double TotalProfit { get; set; }
    }
}
