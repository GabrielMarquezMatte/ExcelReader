using System.Collections;
using System.Diagnostics.CodeAnalysis;
using ExcelReader.Core.Reader;
using ExcelReader.Core.ValueObjects;

namespace ExcelReader.Core.Parser.Internal
{
    [SuppressMessage("Design", "CA1034:Nested types should not be visible",
        Justification = "Public nested struct Enumerator is the standard foreach pattern.")]
    public sealed class ExcelEnumerable<T> : IEnumerable<T> where T : new()
    {
        private readonly XlsxReader _reader;
        private readonly ExcelParserConfig _config;

        internal ExcelEnumerable(XlsxReader reader, ExcelParserConfig config)
        {
            _reader = reader;
            _config = config;
        }

        [SuppressMessage("IDisposableAnalyzers.Correctness", "IDISP015:Member should not return created and cached instance",
            Justification = "Each call creates a fresh enumerator; no caching.")]
        public Enumerator GetEnumerator()
        {
            TypeMapInfo<T> info = TypeMapper<T>.GetInfo();
            XlsxReader.Enumerator rows = _reader.GetEnumerator();
            return new Enumerator(rows, info, _config.ColumnNameComparer, _config.HeaderRow, _reader.IsDate1904);
        }

        IEnumerator<T> IEnumerable<T>.GetEnumerator()
        {
            return GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        public struct Enumerator : IEnumerator<T>
        {
            [SuppressMessage("IDisposableAnalyzers.Correctness", "IDISP006:Implement IDisposable",
                Justification = "Struct implements IDisposable; rows disposed in Dispose().")]
            private XlsxReader.Enumerator _rows;
            private readonly TypeMapInfo<T> _typeInfo;
            private readonly StringComparer _comparer;
            private readonly int _headerRow;
            private readonly bool _isDate1904;
            private int _rowNumber;
            private ColumnParser<T>?[]? _columnParsers;
            private int _maxColumn;
            private T _current = default!;

            internal Enumerator(
                XlsxReader.Enumerator rows,
                TypeMapInfo<T> typeInfo,
                StringComparer comparer,
                int headerRow,
                bool isDate1904)
            {
                _rows = rows;
                _typeInfo = typeInfo;
                _comparer = comparer;
                _headerRow = headerRow;
                _isDate1904 = isDate1904;
            }

            public T Current
            {
                get
                {
                    return _current;
                }
            }

            object? IEnumerator.Current
            {
                get
                {
                    return _current;
                }
            }

            public bool MoveNext()
            {
                while (_rows.MoveNext())
                {
                    _rowNumber++;
                    Row row = _rows.Current;
                    if (_rowNumber < _headerRow)
                    {
                        continue;
                    }
                    if (_rowNumber == _headerRow)
                    {
                        BuildColumnMap(in row);
                        continue;
                    }
                    if (_columnParsers is null)
                    {
                        return false;
                    }
                    _current = new T();
                    ParseCurrentRow(in row);
                    return true;
                }
                return false;
            }

            private void BuildColumnMap(in Row row)
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

            private void ParseCurrentRow(in Row row)
            {
                int bound = Math.Min(_maxColumn, row.ColumnCount);
                for (int col = 0; col < bound; col++)
                {
                    ColumnParser<T>? parser = _columnParsers![col];
                    if (parser is null)
                    {
                        continue;
                    }
                    parser(ref _current, in row, col, _isDate1904);
                }
            }

            [SuppressMessage("IDisposableAnalyzers.Correctness", "IDISP007:Don't dispose injected",
                Justification = "The enumerator owns _rows — it was created by GetEnumerator, not injected from outside.")]
            public void Dispose()
            {
                _rows.Dispose();
            }

            public void Reset()
            {
                throw new NotSupportedException();
            }
        }
    }
}
