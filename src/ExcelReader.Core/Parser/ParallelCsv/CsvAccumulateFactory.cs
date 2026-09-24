using ExcelReader.Core.Reader;
using ExcelReader.Core.Reader.Csv;

namespace ExcelReader.Core.Parser.ParallelCsv
{
    internal delegate CsvRowAction<TState> CsvAccumulateFactory<TState>(Row header, bool sequential);
}
