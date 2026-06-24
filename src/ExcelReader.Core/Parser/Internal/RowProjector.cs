using ExcelReader.Core.ValueObjects;

namespace ExcelReader.Core.Parser.Internal
{
    internal struct RowProjector<T> where T : new()
    {
        private readonly TypeMapInfo<T> _typeInfo;
        private readonly StringComparer _comparer;
        private readonly bool _isDate1904;
        private ColumnParser<T>?[]? _columnParsers;
        private int _maxColumn;

        internal RowProjector(TypeMapInfo<T> typeInfo, StringComparer comparer, bool isDate1904)
        {
            _typeInfo = typeInfo;
            _comparer = comparer;
            _isDate1904 = isDate1904;
        }

        internal readonly bool IsMapped => _columnParsers is not null;

        internal void BuildColumnMap(in Row row)
        {
            _maxColumn = row.ColumnCount;
            _columnParsers = new ColumnParser<T>?[_maxColumn];
            for (int col = 0; col < _maxColumn; col++)
            {
                Cell cell = row[col];
                string header = cell.GetString();
                if (string.IsNullOrEmpty(header))
                {
                    continue;
                }
                int index = _typeInfo.FindIndex(header, _comparer);
                if (index >= 0)
                {
                    _columnParsers[col] = _typeInfo.Parsers[index];
                }
            }
        }

        internal readonly void ParseCurrentRow(in Row row, ref T model)
        {
            int bound = Math.Min(_maxColumn, row.ColumnCount);
            for (int col = 0; col < bound; col++)
            {
                ColumnParser<T>? parser = _columnParsers![col];
                if (parser is null)
                {
                    continue;
                }
                parser(ref model, in row, col, _isDate1904);
            }
        }
    }
}
