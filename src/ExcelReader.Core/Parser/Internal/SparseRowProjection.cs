using ExcelReader.Core.Enums;
using ExcelReader.Core.ValueObjects;

namespace ExcelReader.Core.Parser.Internal
{
    internal static class SparseRowProjection
    {
        internal static ColumnBinding<TModel>[] BuildColumnMap<TModel>(
            in Row row,
            TypeMapInfo<TModel> typeInfo,
            StringComparer comparer,
            HeaderNormalization normalization,
            out int requireValueCount)
            where TModel : allows ref struct
        {
            int propertyCount = typeInfo.PropertyCount;
            int[] columns = new int[propertyCount];
            int[] aliasIndexes = new int[propertyCount];
            var parsers = new ColumnParser<TModel>?[propertyCount];
            Array.Fill(aliasIndexes, int.MaxValue);

            int bindingCount = 0;
            foreach (RowCell rowCell in row.Cells)
            {
                Cell cell = rowCell.Value;
                string header = normalization.Apply(cell.GetString());
                if (string.IsNullOrEmpty(header))
                {
                    continue;
                }
                if (!typeInfo.TryFindHeader(header, comparer, normalization, out HeaderMatch<TModel> match))
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

            typeInfo.ValidateRequiredColumns(aliasIndexes);

            var bindings = new ColumnBinding<TModel>[bindingCount];
            int index = 0;
            int requireValue = 0;
            for (int i = 0; i < parsers.Length; i++)
            {
                ColumnParser<TModel>? parser = parsers[i];
                if (parser is not null)
                {
                    bool requires = typeInfo.RequiresValue(i);
                    if (requires)
                    {
                        requireValue++;
                    }
                    bindings[index++] = new ColumnBinding<TModel>(columns[i], parser, requires, typeInfo.DisplayName(i));
                }
            }
            Array.Sort(bindings, static (left, right) => left.Column.CompareTo(right.Column));
            requireValueCount = requireValue;
            return bindings;
        }

        internal static void ParseRow<TModel>(
            in Row row,
            ColumnBinding<TModel>[] bindings,
            scoped Span<bool> seen,
            bool track,
            bool isDate1904,
            IFormatProvider provider,
            bool throwOnParseFailure,
            int rowNumber,
            ref TModel model)
            where TModel : allows ref struct
        {
            if (track)
            {
                seen[..bindings.Length].Clear();
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
                ref readonly ColumnBinding<TModel> binding = ref bindings[bindingIndex];
                if (binding.Column != column)
                {
                    continue;
                }
                Cell cell = rowCell.Value;
                if (cell.Type == CellType.Empty)
                {
                    bindingIndex++;
                    continue;
                }
                if (binding.Parser(ref model, in cell, isDate1904, provider))
                {
                    if (track && binding.RequireValue)
                    {
                        seen[bindingIndex] = true;
                    }
                }
                else if (throwOnParseFailure)
                {
                    throw new ExcelParseException(rowNumber, binding.Name, cell.GetString());
                }
                bindingIndex++;
            }
            if (track)
            {
                ValidateRowValues(bindings, seen, rowNumber);
            }
        }

        private static void ValidateRowValues<TModel>(ColumnBinding<TModel>[] bindings, ReadOnlySpan<bool> seen, int rowNumber)
            where TModel : allows ref struct
        {
            for (int i = 0; i < bindings.Length; i++)
            {
                if (bindings[i].RequireValue && !seen[i])
                {
                    throw ProjectionRules.MissingRequiredValue(bindings[i].Name, rowNumber);
                }
            }
        }
    }
}
