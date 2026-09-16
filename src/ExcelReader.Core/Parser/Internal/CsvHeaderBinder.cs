using ExcelReader.Core.Reader;

namespace ExcelReader.Core.Parser.Internal
{
    internal static class CsvHeaderBinder
    {
        internal static CsvBoundColumnMap<T> Bind<T>(
            CsvReader reader,
            ExcelParserConfig config,
            TypeMapInfo<T> info,
            out long firstDataRecordOffset)
        {
            using CsvReader.Enumerator rows = reader.GetEnumerator();
            int rowNumber = 0;
            while (rows.MoveNext())
            {
                rowNumber++;
                if (rowNumber != config.HeaderRow)
                {
                    continue;
                }
                CsvBoundColumnMap<T> map = CsvRowProjector<T>.BuildBoundMap(
                    rows, info, config.ColumnNameComparer, config.HeaderNormalization);
                firstDataRecordOffset = rows.MoveNext() ? rows.CurrentRecordStart : long.MaxValue;
                return map;
            }
            throw new InvalidOperationException(
                $"The CSV source has no row at header index {config.HeaderRow}.");
        }
    }
}
