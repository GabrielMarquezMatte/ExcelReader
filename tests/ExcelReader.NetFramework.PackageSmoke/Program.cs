using System;
using ExcelReader.NetFramework.Parser;
using ExcelReader.NetFramework.Reader;
using ExcelReader.NetFramework.Writer;

namespace PackageSmoke
{
    public sealed class Row
    {
        [ExcelColumn("Coluna1")]
        public string? Coluna1 { get; set; }
    }

    internal static class Program
    {
        private static int Main(string[] args)
        {
            using ExcelWorkbook workbook = ExcelWorkbook.Open(args[0]);
            Row[] rows = workbook.Sheet(0).Read<Row>();
            Console.WriteLine(rows.Length);
            if (rows.Length != 100)
            {
                return 1;
            }

            if (rows[0].Coluna1 != "Valor1")
            {
                return 2;
            }

            using ExcelWorkbookWriter writer = ExcelWorkbookWriter.CreateInMemory(ExcelFormat.Xlsx);
            writer.WriteSheet("Copy", rows);
            byte[] bytes = writer.CloseToArray();
            using ExcelWorkbook copy = ExcelWorkbook.Open(bytes);
            Row[] copied = copy.Sheet("Copy").Read<Row>();
            return copied.Length == 100 && copied[0].Coluna1 == "Valor1" ? 0 : 3;
        }
    }
}
