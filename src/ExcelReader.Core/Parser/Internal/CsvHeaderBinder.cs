using ExcelReader.Core.Reader;

namespace ExcelReader.Core.Parser.Internal
{
    // Reads a CSV source's header exactly once, on one thread, and produces the shareable column map
    // every parallel worker projects through. Workers never see a header: chunk 0 begins at
    // firstDataRecordOffset, and every later chunk begins mid-file by construction.
    internal static class CsvHeaderBinder
    {
        // Returns the bound map, and reports through firstDataRecordOffset the byte offset of the
        // first data record — the offset chunk 0 starts at.
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
                // 1-based, matching ProjectionRules.ClassifyRow's convention (rowNumber is
                // incremented before comparison, so the first record read is row 1).
                rowNumber++;
                if (rowNumber != config.HeaderRow)
                {
                    continue;
                }
                CsvBoundColumnMap<T> map = CsvRowProjector<T>.BuildBoundMap(
                    rows, info, config.ColumnNameComparer, config.HeaderNormalization);
                // The header record has been consumed; the next MoveNext lands on the first data
                // record, whose offset is where chunk 0 must begin.
                firstDataRecordOffset = rows.MoveNext() ? rows.CurrentRecordStart : long.MaxValue;
                return map;
            }
            throw new InvalidOperationException(
                $"The CSV source has no row at header index {config.HeaderRow}.");
        }
    }
}
