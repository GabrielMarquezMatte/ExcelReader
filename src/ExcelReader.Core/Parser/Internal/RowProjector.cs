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
        private readonly ExcelRowContext _context;
        private readonly bool _throwOnParseFailure;
        private readonly bool _skipBlankRows;
        private readonly bool _skipEmptyRows;
        private ColumnBinding<T>[]? _bindings;
        private bool[] _seen;
        private int _requireValueCount;
        private int _rowNumber;

        internal RowProjector(TypeMapInfo<T> typeInfo, StringComparer comparer, HeaderNormalization normalization, int headerRow, bool isDate1904, IFormatProvider provider, bool throwOnParseFailure = false, bool skipBlankRows = false, bool skipEmptyRows = false)
        {
            _typeInfo = typeInfo;
            _comparer = comparer;
            _normalization = normalization;
            _headerRow = headerRow;
            _context = new ExcelRowContext { IsDate1904 = isDate1904, FormatProvider = provider };
            _throwOnParseFailure = throwOnParseFailure;
            _skipBlankRows = skipBlankRows;
            _skipEmptyRows = skipEmptyRows;
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
                return YieldUnlessBlank(in row);
            }
            ProjectionStep step = ProjectionRules.ClassifyRow(ref _rowNumber, _headerRow, _bindings is not null);
            if (step == ProjectionStep.Yield)
            {
                return YieldUnlessBlank(in row);
            }
            if (step != ProjectionStep.BuildMap)
            {
                return step;
            }
            BuildColumnMap(in row);
            return ProjectionStep.Skip;
        }

        // A CSV blank line is a record of one empty field; the CSV-specialized projector skips it the same way.
        private readonly ProjectionStep YieldUnlessBlank(in Row row)
        {
            if (_skipBlankRows && row.Descriptors is [{ Type: CellType.Empty }])
            {
                return ProjectionStep.Skip;
            }
            if (_skipEmptyRows && !SparseRowProjection.HasBoundValue(in row, _bindings!))
            {
                return ProjectionStep.Skip;
            }
            return ProjectionStep.Yield;
        }

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
            SparseRowProjection.ParseRow(in row, _bindings!, _seen, track, in _context, _throwOnParseFailure, _rowNumber, ref model);
        }
    }
}
