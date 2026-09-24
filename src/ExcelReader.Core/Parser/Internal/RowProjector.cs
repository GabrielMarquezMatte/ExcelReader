using System.Runtime.CompilerServices;
using ExcelReader.Core.Reader;


namespace ExcelReader.Core.Parser.Internal
{
    internal struct RowProjector<T>
        where T : allows ref struct
    {
        private readonly TypeMapInfo<T> _typeInfo;
        private readonly StringComparer _comparer;
        private readonly HeaderNormalization _normalization;
        private readonly int _headerRow;
        private readonly bool _isDate1904;
        private readonly IFormatProvider _provider;
        private readonly bool _throwOnParseFailure;
        private ColumnBinding<T>[]? _bindings;
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

        internal ProjectionStep Classify(in Row row)
        {
            if (_typeInfo.IsIndexBased)
            {
                if (_bindings is null)
                {
                    BuildIndexColumnMap();
                }
                _rowNumber++;
                return ProjectionStep.Yield;
            }
            ProjectionStep step = ProjectionRules.ClassifyRow(ref _rowNumber, _headerRow, _bindings is not null);
            if (step != ProjectionStep.BuildMap)
            {
                return step;
            }
            BuildColumnMap(in row);
            return ProjectionStep.Skip;
        }

        // Kept out of line so the caller's foreach tier-1 compile can't inline the whole projection and spill its hot loop.
        [MethodImpl(MethodImplOptions.NoInlining)]
        internal readonly T Project(Row row)
        {
            T model = _typeInfo.CreateInstance();
            ParseCurrentRow(in row, ref model);
            return model;
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
