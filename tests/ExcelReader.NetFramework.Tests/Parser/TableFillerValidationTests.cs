using System.Runtime.InteropServices;
using ExcelReader.NetFramework.Interop;
using ExcelReader.NetFramework.Parser.Internal;
using ExcelReader.NetFramework.Reader;

namespace ExcelReader.NetFramework.Tests.Parser
{
    public sealed unsafe class TableFillerValidationTests
    {
        [Fact]
        public void ColumnLengthMismatchIsRejected()
        {
            ExcelReaderException ex = Assert.Throws<ExcelReaderException>(() => FillWith(Xl.XL_T_I64, 5));
            Assert.Equal(ExcelStatus.Error, ex.Status);
            Assert.Contains("Coluna3", ex.Message, StringComparison.Ordinal);
            Assert.Contains("returned 5 rows", ex.Message, StringComparison.Ordinal);
            Assert.Contains("expected 3.", ex.Message, StringComparison.Ordinal);
        }

        [Fact]
        public void ColumnTypeMismatchIsRejected()
        {
            ExcelReaderException ex = Assert.Throws<ExcelReaderException>(() => FillWith(Xl.XL_T_STRING, 3));
            Assert.Equal(ExcelStatus.Error, ex.Status);
            Assert.Contains("Coluna3", ex.Message, StringComparison.Ordinal);
            Assert.Contains("column type 0 for", ex.Message, StringComparison.Ordinal);
            Assert.Contains("expected 1.", ex.Message, StringComparison.Ordinal);
        }

        private static void FillWith(int columnType, long columnLength)
        {
            PropertyBinder<Row> binder = TypeMap<Row>.Instance.Properties.Single(p => string.Equals(p.PropertyName, "Coluna3", StringComparison.Ordinal));
            IntPtr memory = Marshal.AllocHGlobal(sizeof(XlColumn));
            try
            {
                XlColumn* column = (XlColumn*)memory;
                *column = new XlColumn { Type = columnType, Length = columnLength };
                XlTable table = new XlTable { ColumnCount = 1, RowCount = 3, Columns = column };
                Row[] items = { new Row(), new Row(), new Row() };
                TableFiller.Fill(new[] { binder }, in table, items, 0, 0);
            }
            finally
            {
                Marshal.FreeHGlobal(memory);
            }
        }
    }
}
