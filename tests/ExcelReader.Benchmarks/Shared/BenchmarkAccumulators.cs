using System.Data;
using ExcelReader.Core.Reader;
using Sylvan.Data.Excel;

namespace ExcelReader.Benchmarks
{
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

        internal static long AccumulateRowMaterialized(Row row)
        {
            long acc = 0;
            foreach (RowCell rowCell in row.Cells)
            {
                Cell cell = rowCell.Value;
                switch (cell.Type)
                {
                    case CellType.ExcelString:
                        acc += cell.GetString().Length;
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

        internal static long AccumulateSylvanExcel(Sylvan.Data.Excel.ExcelDataReader reader)
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

        internal static long AccumulateDataReader(IDataReader reader)
        {
            long acc = 0;
            while (reader.Read())
            {
                for (int i = 0; i < reader.FieldCount; i++)
                {
                    switch (reader.GetValue(i))
                    {
                        case string s: acc += s.Length; break;
                        case double d: acc += (long)d; break;
                        case DateTime dt: acc += dt.Ticks; break;
                    }
                }
            }
            return acc;
        }
    }
}
