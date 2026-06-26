using ExcelReader.Core.Enums;
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
        private ColumnBinding<T>[]? _bindings;
        // Per-row scratch: _seen[i] is set when binding i saw a non-empty cell this row. Only allocated
        // and walked when at least one bound column requires a value (_requireValueCount > 0).
        private bool[]? _seen;
        private int _requireValueCount;
        private int _rowNumber;

        internal RowProjector(TypeMapInfo<T> typeInfo, StringComparer comparer, HeaderNormalization normalization, int headerRow, bool isDate1904, IFormatProvider provider)
        {
            _typeInfo = typeInfo;
            _comparer = comparer;
            _normalization = normalization;
            _headerRow = headerRow;
            _isDate1904 = isDate1904;
            _provider = provider;
        }

        // The per-row state machine shared by every enumerator (sync/async, xlsx/xls): skip rows before
        // the header, build the column map at the header row, then project each subsequent row.
        internal ProjectionStep Advance(in Row row, ref T model)
        {
            _rowNumber++;
            if (_rowNumber < _headerRow)
            {
                return ProjectionStep.Skip;
            }
            if (_rowNumber == _headerRow)
            {
                BuildColumnMap(in row);
                return ProjectionStep.Skip;
            }
            if (_bindings is null)
            {
                return ProjectionStep.Stop;
            }
            model = _typeInfo.CreateInstance();
            ParseCurrentRow(in row, ref model);
            return ProjectionStep.Yield;
        }

        private void BuildColumnMap(in Row row)
        {
            int propertyCount = _typeInfo.PropertyCount;
            int[] columns = new int[propertyCount];
            int[] aliasIndexes = new int[propertyCount];
            var parsers = new ColumnParser<T>?[propertyCount];
            Array.Fill(aliasIndexes, int.MaxValue);

            int bindingCount = 0;
            foreach (RowCell rowCell in row.Cells)
            {
                Cell cell = rowCell.Value;
                string header = _normalization.Apply(cell.GetString());
                if (string.IsNullOrEmpty(header))
                {
                    continue;
                }
                if (!_typeInfo.TryFindHeader(header, _comparer, _normalization, out HeaderMatch<T> match))
                {
                    continue;
                }
                if (match.AliasIndex >= aliasIndexes[match.PropertyIndex])
                {
                    continue;
                }
                if (aliasIndexes[match.PropertyIndex] == int.MaxValue)
                {
                    bindingCount++;
                }
                columns[match.PropertyIndex] = rowCell.ColumnIndex;
                parsers[match.PropertyIndex] = match.Parser;
                aliasIndexes[match.PropertyIndex] = match.AliasIndex;
            }

            _typeInfo.ValidateRequiredColumns(aliasIndexes);

            var bindings = new ColumnBinding<T>[bindingCount];
            int index = 0;
            int requireValueCount = 0;
            for (int i = 0; i < parsers.Length; i++)
            {
                ColumnParser<T>? parser = parsers[i];
                if (parser is not null)
                {
                    bool requireValue = _typeInfo.RequiresValue(i);
                    if (requireValue)
                    {
                        requireValueCount++;
                    }
                    bindings[index++] = new ColumnBinding<T>(columns[i], parser, requireValue, _typeInfo.DisplayName(i));
                }
            }
            Array.Sort(bindings, static (left, right) => left.Column.CompareTo(right.Column));
            _bindings = bindings;
            _requireValueCount = requireValueCount;
            _seen = requireValueCount > 0 ? new bool[bindings.Length] : null;
        }

        private readonly void ParseCurrentRow(in Row row, ref T model)
        {
            ColumnBinding<T>[] bindings = _bindings!;
            bool track = _requireValueCount > 0;
            if (track)
            {
                Array.Clear(_seen!, 0, bindings.Length);
            }
            int bindingIndex = 0;
            foreach (RowCell rowCell in row.Cells)
            {
                int column = rowCell.ColumnIndex;
                while (bindingIndex < bindings.Length && bindings[bindingIndex].Column < column)
                {
                    bindingIndex++;
                }
                if (bindingIndex == bindings.Length)
                {
                    break;
                }
                ColumnBinding<T> binding = bindings[bindingIndex];
                if (binding.Column == column)
                {
                    Cell cell = rowCell.Value;
                    binding.Parser(ref model, in cell, _isDate1904, _provider);
                    if (track && binding.RequireValue && cell.Type != CellType.Empty)
                    {
                        _seen![bindingIndex] = true;
                    }
                    bindingIndex++;
                }
            }
            if (track)
            {
                ValidateRowValues(bindings);
            }
        }

        // Throws on the first required column whose cell was empty or absent in the current row.
        private readonly void ValidateRowValues(ColumnBinding<T>[] bindings)
        {
            for (int i = 0; i < bindings.Length; i++)
            {
                if (bindings[i].RequireValue && !_seen![i])
                {
                    throw new InvalidOperationException(
                        $"Required column '{bindings[i].Name}' has no value in row {_rowNumber.ToString(System.Globalization.CultureInfo.InvariantCulture)}.");
                }
            }
        }

        private readonly struct ColumnBinding<TModel>
        {
            internal ColumnBinding(int column, ColumnParser<TModel> parser, bool requireValue, string name)
            {
                Column = column;
                Parser = parser;
                RequireValue = requireValue;
                Name = name;
            }

            internal int Column { get; }
            internal ColumnParser<TModel> Parser { get; }
            internal bool RequireValue { get; }
            internal string Name { get; }
        }
    }
}
