using ExcelReader.Core.Enums;
using ExcelReader.Core.ValueObjects;

namespace ExcelReader.Core.Parser.Internal
{
    // The merge-walk column-binding loop shared by RowProjector<T> (class/struct models) and
    // NamedRefRowEnumerator<TModel,...> (ref struct models, net9+). Both bind sparse Row.Cells to a
    // header-resolved column map the same way; before this they carried two byte-identical copies (see
    // docs/road-to-a.md F9). A `static` generic method — never storing TModel in a field — is what lets
    // a ref-struct-constrained TModel flow through without CS8345.
    //
    // Deliberately NOT shared with CsvRowProjector<T>: CSV rows are dense (field index == column index,
    // no gaps), so its fast path is a direct indexed scan with no merge-walk at all — forcing it through
    // this sparse-cursor shape would trade a genuine O(1)-per-field win for unification that saves no
    // real duplication (its BuildColumnMap/ParseCurrentRow are structurally different, not copies of
    // this one). CsvRowProjector applies the same F3 parse-failure policy inline instead.
    internal static class SparseRowProjection
    {
        internal static ColumnBinding<TModel>[] BuildColumnMap<TModel>(
            in Row row,
            TypeMapInfo<TModel> typeInfo,
            StringComparer comparer,
            HeaderNormalization normalization,
            out int requireValueCount)
#if NET9_0_OR_GREATER
            where TModel : allows ref struct
#endif
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

        // On a parse failure (non-empty cell, parser returned false): throws ExcelParseException when
        // throwOnParseFailure is set; otherwise leaves `seen` unset for that binding, so a [ExcelRequired]
        // column with an unparseable value fails the same way as a blank one via ValidateRowValues below
        // (F3) — closing the "required-but-unparseable silently passes" gap even with the flag off.
        internal static void ParseRow<TModel>(
            in Row row,
            ColumnBinding<TModel>[] bindings,
            bool[] seen,
            bool track,
            bool isDate1904,
            IFormatProvider provider,
            bool throwOnParseFailure,
            int rowNumber,
            ref TModel model)
#if NET9_0_OR_GREATER
            where TModel : allows ref struct
#endif
        {
            if (track)
            {
                Array.Clear(seen, 0, bindings.Length);
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

        private static void ValidateRowValues<TModel>(ColumnBinding<TModel>[] bindings, bool[] seen, int rowNumber)
#if NET9_0_OR_GREATER
            where TModel : allows ref struct
#endif
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
