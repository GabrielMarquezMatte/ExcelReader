using ExcelReader.Core.Enums;
using ExcelReader.Core.ValueObjects;
using Sylvan.Data.Excel;

namespace ExcelReader.Benchmarks
{
    // Shared consume loops for the cell-by-cell read benchmarks. Kept in one place so the
    // ExcelReader and Sylvan sides stay comparable across every benchmark class that uses them —
    // a CellType case added to only one copy would silently skew the comparison.
    internal static class BenchmarkAccumulators
    {
        internal static long AccumulateRow(Row row)
        {
            long acc = 0;
            foreach (RowCell rowCell in row.Cells)
            {
                Cell cell = rowCell.Value;
                switch (cell.Type)
                {
                    case CellType.ExcelString:
                        acc += cell.Value.Length;
                        break;
                    case CellType.Number:
                        if (cell.TryParse(null, out double n)) { acc += (long)n; }
                        break;
                    case CellType.Date:
                        if (cell.TryGetDateTime(out DateTime d)) { acc += d.Ticks; }
                        break;
                }
            }
            return acc;
        }

        internal static long AccumulateSylvanExcel(ExcelDataReader reader)
        {
            long acc = 0;
            do
            {
                while (reader.Read())
                {
                    for (int i = 0; i < reader.FieldCount; i++)
                    {
                        if (reader.IsDBNull(i)) { continue; }
                        switch (reader.GetExcelDataType(i))
                        {
                            case ExcelDataType.String:
                                acc += reader.GetString(i).Length;
                                break;
                            case ExcelDataType.Numeric:
                                acc += (long)reader.GetDouble(i);
                                break;
                            case ExcelDataType.DateTime:
                                acc += reader.GetDateTime(i).Ticks;
                                break;
                        }
                    }
                }
            }
            while (reader.NextResult());
            return acc;
        }
    }
}
