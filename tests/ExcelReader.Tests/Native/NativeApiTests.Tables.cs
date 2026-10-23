using System.Text;
using ExcelReader.Core.Reader;
using ExcelReader.Core.Writer;
using ExcelReader.Native;
using ExcelReader.Native.Reading;
using ExcelReader.Native.Writer;

namespace ExcelReader.Tests.Native
{
    public sealed partial class NativeApiTests
    {
        private static byte[] WriteSalesTable(int format, int degreeOfParallelism, NativeExcelTable excelTable)
        {
            NativeTable sales = BuildQtyNameTable([3L, 7L], ["widget", "gadget"]);
            try
            {
                NativeWriteOptionsRaw raw = DefaultWriteOptionsRaw();
                raw.DegreeOfParallelism = degreeOfParallelism;
                raw.BatchRows = 1;
                Assert.True(NativeWriteOptions.TryDecode(raw, null, out NativeWriteOptions options, out _));
                Assert.Equal(NativeStatus.Ok, WriteApi.OpenWriteHandleToMemory(format, options, out NativeWriterHandle? handle));
                Assert.NotNull(handle);
                using (handle)
                {
                    handle.StartSheet("Notas");
                    handle.StartRow();
                    handle.WriteString("antes");
                    handle.EndRow();
                    handle.EndSheet();
                    handle.WriteSheet("Vendas", QtyNameSpecs(), sales, hasHeader: true, excelTable);
                    Assert.Equal(NativeStatus.Ok, WriteApi.GetWriteHandleBytes(handle, out byte[]? bytes));
                    return bytes!;
                }
            }
            finally
            {
                FreeBuiltTable(ref sales);
            }
        }

        private static List<string> ReadViewRows(NativeHandle handle, int sheet)
        {
            List<string> rows = [];
            using NativeRowCursor cursor = OpenRows(handle, sheet);
            while (ReadApi.NextRowView(cursor, out NativeRow row) == NativeStatus.Ok)
            {
                rows.Add(string.Join('|', ReadView(row).Select(cell => cell.Value)));
            }
            return rows;
        }

        [Theory]
        [InlineData(NativeFormat.Xlsx, 1)]
        [InlineData(NativeFormat.Xlsb, 1)]
        [InlineData(NativeFormat.Xlsx, 4)]
        [InlineData(NativeFormat.Xlsb, 4)]
        public void WriteSheet_Should_Write_An_Excel_Table_The_Table_Exports_Read_Back(int format, int degreeOfParallelism)
        {
            var excelTable = new NativeExcelTable("Sales", ExcelTableOptions.Default with { StyleName = "TableStyleLight9", ShowColumnStripes = true });
            byte[] bytes = WriteSalesTable(format, degreeOfParallelism, excelTable);

            using (IExcelWorkbook workbook = Excel.Open(bytes.AsMemory()))
            {
                ExcelTable written = Assert.Single(workbook.Tables);
                Assert.Equal("TableStyleLight9", written.StyleName);
                Assert.Equal(["qty", "name"], written.ColumnNames, StringComparer.Ordinal);
            }

            Assert.Equal(NativeStatus.Ok, ReadApi.OpenMemory(bytes, format, out NativeHandle? handle));
            try
            {
                Assert.Equal(NativeStatus.Ok, ReadApi.TableCount(handle, out int count));
                Assert.Equal(1, count);
                Assert.Equal(NativeStatus.Ok, ReadApi.TableInfoAt(handle, 0, out NativeTableInfo info));
                Assert.Equal((1, 0, 2, 0, 1, 1, 0),
                    (info.Sheet, info.FirstRow, info.LastRow, info.FirstColumn, info.LastColumn, info.HeaderRowCount, info.TotalsRowCount));

                Span<byte> name = stackalloc byte[16];
                Assert.Equal(NativeStatus.Ok, ReadApi.TableNameAt(handle, 0, name, out int length));
                Assert.Equal("Sales", Encoding.UTF8.GetString(name[..length]));
                Assert.Equal(NativeStatus.BufferTooSmall, ReadApi.TableNameAt(handle, 0, name[..2], out int needed));
                Assert.Equal(5, needed);

                Assert.Equal(NativeStatus.Ok, ReadApi.TableIndex(handle, "sales"u8, out int index));
                Assert.Equal(0, index);
                Assert.Equal(NativeStatus.Ok, ReadApi.TableIndex(handle, "missing"u8, out int missing));
                Assert.Equal(-1, missing);

                Assert.Equal(["qty|name", "3|widget", "7|gadget"], ReadViewRows(handle!, ReadApi.FirstTableSheet), StringComparer.Ordinal);
                Assert.Equal(NativeStatus.Ok, ReadApi.SheetNameAt(handle, ReadApi.FirstTableSheet, name, out length));
                Assert.Equal("Vendas", Encoding.UTF8.GetString(name[..length]));
            }
            finally
            {
                ReadApi.Close(handle);
            }
        }

        [Fact]
        public void Table_Exports_Should_Reject_A_Table_That_Does_Not_Exist()
        {
            byte[] bytes = WriteSalesTable(NativeFormat.Xlsx, 1, new NativeExcelTable("Sales", ExcelTableOptions.Default));
            Assert.Equal(NativeStatus.Ok, ReadApi.OpenMemory(bytes, NativeFormat.Xlsx, out NativeHandle? handle));
            try
            {
                Assert.Equal(NativeStatus.Error, ReadApi.TableInfoAt(handle, 1, out _));
                Assert.Equal(NativeStatus.InvalidArgument, ReadApi.TableInfoAt(handle, -1, out _));
                Assert.Equal(NativeStatus.Error, ReadApi.OpenRows(handle, ReadApi.FirstTableSheet - 1, out _));
                Assert.Equal(NativeStatus.InvalidArgument, ReadApi.OpenRows(handle, -1, out _));
            }
            finally
            {
                ReadApi.Close(handle);
            }
        }

        [Theory]
        [InlineData(NativeFormat.Xls)]
        [InlineData(NativeFormat.Csv)]
        public void WriteSheet_Should_Refuse_An_Excel_Table_Outside_Xlsx_And_Xlsb(int format)
        {
            NativeTable sales = BuildQtyNameTable([3L], ["widget"]);
            try
            {
                Assert.Equal(NativeStatus.Ok, WriteApi.OpenWriteHandleToMemory(format, DefaultWriteOptions(), out NativeWriterHandle? handle));
                Assert.NotNull(handle);
                using (handle)
                {
                    Assert.Throws<NotSupportedException>(() =>
                        handle.WriteSheet("Vendas", QtyNameSpecs(), sales, hasHeader: true, new NativeExcelTable("Sales", ExcelTableOptions.Default)));
                    handle.WriteSheet("Vendas", QtyNameSpecs(), sales, hasHeader: true, excelTable: null);
                }
            }
            finally
            {
                FreeBuiltTable(ref sales);
            }
        }

        [Fact]
        public unsafe void SheetOptions_Should_Decode_A_Table_And_Its_Flags()
        {
            fixed (byte* name = "Sales"u8, style = "TableStyleDark1"u8)
            {
                var raw = new NativeSheetOptionsRaw
                {
                    StructSize = sizeof(NativeSheetOptionsRaw),
                    TableName = name,
                    TableNameLen = 5,
                    TableStyle = style,
                    TableStyleLen = 15,
                    TableFlags = NativeExcelTable.NoRowStripes | NativeExcelTable.FirstColumn,
                };
                Assert.True(NativeExcelTable.TryRead(&raw, out NativeExcelTable? table, out _));
                Assert.Equal(new NativeExcelTable("Sales", new ExcelTableOptions
                {
                    StyleName = "TableStyleDark1",
                    ShowRowStripes = false,
                    ShowFirstColumn = true,
                }), table);

                raw.TableStyle = null;
                raw.TableFlags = NativeExcelTable.NoStyle;
                Assert.True(NativeExcelTable.TryRead(&raw, out table, out _));
                Assert.Null(table!.Options.StyleName);

                Assert.True(NativeExcelTable.TryRead(null, out table, out _));
                Assert.Null(table);
            }
        }

        [Fact]
        public unsafe void SheetOptions_Should_Reject_Contradictory_Or_Unknown_Settings()
        {
            fixed (byte* name = "Sales"u8, style = "TableStyleDark1"u8)
            {
                NativeSheetOptionsRaw valid = new() { StructSize = sizeof(NativeSheetOptionsRaw), TableName = name, TableNameLen = 5 };
                NativeSheetOptionsRaw[] invalid =
                [
                    valid with { StructSize = 8 },
                    valid with { TableFlags = 64 },
                    valid with { TableStyle = style, TableStyleLen = 15, TableFlags = NativeExcelTable.NoStyle },
                    valid with { TableName = null, TableFlags = NativeExcelTable.ColumnStripes },
                    valid with { TableNameLen = -1 },
                ];
                foreach (NativeSheetOptionsRaw raw in invalid)
                {
                    NativeSheetOptionsRaw copy = raw;
                    Assert.False(NativeExcelTable.TryRead(&copy, out _, out string? error));
                    Assert.NotEmpty(error);
                }
            }
        }
    }
}
