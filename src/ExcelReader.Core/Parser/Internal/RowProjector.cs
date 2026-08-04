using ExcelReader.Core.ValueObjects;

namespace ExcelReader.Core.Parser.Internal
{
    internal struct RowProjector<T>
    {
        private readonly TypeMapInfo<T> _typeInfo;
        private readonly StringComparer _comparer;
        private readonly HeaderNormalization _normalization;
        private readonly int _headerRow;
        private readonly bool _isDate1904;
        private readonly IFormatProvider _provider;
        private readonly bool _throwOnParseFailure;
        private ColumnBinding<T>[]? _bindings;
        // Per-row scratch: _seen[i] is set when binding i saw a non-empty cell this row that parsed
        // successfully. Only allocated and walked when at least one bound column requires a value
        // (_requireValueCount > 0).
        private bool[] _seen;
        private int _requireValueCount;
        private int _rowNumber;

        internal RowProjector(TypeMapInfo<T> typeInfo, StringComparer comparer, HeaderNormalization normalization, int headerRow, bool isDate1904, IFormatProvider provider, bool throwOnParseFailure = false)
        {
            _typeInfo = typeInfo;
            _comparer = comparer;
            _normalization = normalization;
            _headerRow = headerRow;
            _isDate1904 = isDate1904;
            _provider = provider;
            _throwOnParseFailure = throwOnParseFailure;
            _seen = [];
        }

        // The per-row state machine shared by every enumerator (sync/async, xlsx/xls): skip rows before
        // the header, build the column map at the header row, then project each subsequent row.
        internal ProjectionStep Advance(in Row row, ref T model)
        {
            if (_typeInfo.IsIndexBased)
            {
                if (_bindings is null)
                {
                    BuildIndexColumnMap();
                }
                _rowNumber++;
                model = _typeInfo.CreateInstance();
                ParseCurrentRow(in row, ref model);
                return ProjectionStep.Yield;
            }
            ProjectionStep step = ProjectionRules.ClassifyRow(ref _rowNumber, _headerRow, _bindings is not null);
            if (step == ProjectionStep.BuildMap)
            {
                BuildColumnMap(in row);
                return ProjectionStep.Skip;
            }
            if (step != ProjectionStep.Yield)
            {
                return step;
            }
            model = _typeInfo.CreateInstance();
            ParseCurrentRow(in row, ref model);
            return ProjectionStep.Yield;
        }

        private void BuildColumnMap(in Row row)
        {
            _bindings = SparseRowProjection.BuildColumnMap(in row, _typeInfo, _comparer, _normalization, out int requireValueCount);
            _requireValueCount = requireValueCount;
            _seen = requireValueCount > 0 ? new bool[_bindings.Length] : [];
        }

        private void BuildIndexColumnMap()
        {
            _bindings = _typeInfo.IndexBindings;
            int requireValueCount = 0;
            foreach (ColumnBinding<T> binding in _bindings)
            {
                if (binding.RequireValue)
                {
                    requireValueCount++;
                }
            }
            _requireValueCount = requireValueCount;
            _seen = requireValueCount > 0 ? new bool[_bindings.Length] : [];
        }

        private readonly void ParseCurrentRow(in Row row, ref T model)
        {
            bool track = _requireValueCount > 0;
            SparseRowProjection.ParseRow(in row, _bindings!, _seen, track, _isDate1904, _provider, _throwOnParseFailure, _rowNumber, ref model);
        }
    }
}
