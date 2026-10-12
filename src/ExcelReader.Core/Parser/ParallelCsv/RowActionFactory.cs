using ExcelReader.Core.Reader;
using ExcelReader.Core.Reader.Csv;

namespace ExcelReader.Core.Parser.ParallelCsv
{
    internal delegate RowAction<TState> RowActionFactory<TState>(Row header, bool sequential);
}
