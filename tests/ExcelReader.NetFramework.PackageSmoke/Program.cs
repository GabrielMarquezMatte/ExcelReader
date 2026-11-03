using System;
using System.Collections.Generic;
using ExcelReader.NetFramework;

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
            List<Row> rows = workbook.Sheet(0).Read<Row>();
            Console.WriteLine(rows.Count);
            if (rows.Count != 100)
            {
                return 1;
            }

            return rows[0].Coluna1 == "Valor1" ? 0 : 2;
        }
    }
}
