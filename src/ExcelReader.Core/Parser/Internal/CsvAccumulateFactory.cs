using ExcelReader.Core.Reader;
using ExcelReader.Core.ValueObjects;

namespace ExcelReader.Core.Parser.Internal
{
    internal delegate CsvRowAction<TState> CsvAccumulateFactory<TState>(Row header, bool sequential);
}
