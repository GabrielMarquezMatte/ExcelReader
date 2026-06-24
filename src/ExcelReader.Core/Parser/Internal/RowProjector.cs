using ExcelReader.Core.ValueObjects;

namespace ExcelReader.Core.Parser.Internal
{
    internal struct RowProjector<T> where T : new()
    {
        private readonly TypeMapInfo<T> _typeInfo;
        private readonly StringComparer _comparer;
        private readonly bool _isDate1904;
        private ColumnBinding<T>[]? _bindings;

        internal RowProjector(TypeMapInfo<T> typeInfo, StringComparer comparer, bool isDate1904)
        {
            _typeInfo = typeInfo;
            _comparer = comparer;
            _isDate1904 = isDate1904;
        }

        internal readonly bool IsMapped => _bindings is not null;

        internal void BuildColumnMap(in Row row)
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
                string header = cell.GetString();
                if (string.IsNullOrEmpty(header))
                {
                    continue;
                }
                if (!_typeInfo.TryFindHeader(header, _comparer, out HeaderMatch<T> match))
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

            var bindings = new ColumnBinding<T>[bindingCount];
            int index = 0;
            for (int i = 0; i < parsers.Length; i++)
            {
                ColumnParser<T>? parser = parsers[i];
                if (parser is not null)
                {
                    bindings[index++] = new ColumnBinding<T>(columns[i], parser);
                }
            }
            Array.Sort(bindings, static (left, right) => left.Column.CompareTo(right.Column));
            _bindings = bindings;
        }

        internal readonly void ParseCurrentRow(in Row row, ref T model)
        {
            ColumnBinding<T>[] bindings = _bindings!;
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
                    return;
                }
                ColumnBinding<T> binding = bindings[bindingIndex];
                if (binding.Column == column)
                {
                    Cell cell = rowCell.Value;
                    binding.Parser(ref model, in cell, _isDate1904);
                    bindingIndex++;
                }
            }
        }

        private readonly struct ColumnBinding<TModel>
            where TModel : new()
        {
            internal ColumnBinding(int column, ColumnParser<TModel> parser)
            {
                Column = column;
                Parser = parser;
            }

            internal int Column { get; }
            internal ColumnParser<TModel> Parser { get; }
        }
    }
}
