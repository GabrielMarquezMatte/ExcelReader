using ExcelReader.Core.Reader;
using ExcelReader.Core.ValueObjects;

namespace ExcelReader.Core.Parser.Internal
{
    internal static class MappedProjection<TState, TModel>
        where TModel : allows ref struct
    {
        private const int StackallocLimit = 256;

        internal static CsvRowAction<TState> Bind(
            CsvModelMap<TModel> map, Row header, int headerRow, bool sequential, CsvModelSink<TState, TModel> sink)
        {
            TypeMapInfo<TModel> info = map.Info;
            ExcelParserConfig config = map.Config;
            ColumnBinding<TModel>[] bindings;
            int requireValueCount;
            if (info.IsIndexBased)
            {
                bindings = info.IndexBindings;
                requireValueCount = bindings.Count(static b => b.RequireValue);
            }
            else
            {
                bindings = SparseRowProjection.BuildColumnMap(in header, info, config.ColumnNameComparer, config.HeaderNormalization, out requireValueCount);
            }
            bool track = requireValueCount > 0;
            IFormatProvider provider = config.Culture;
            bool throwOnParseFailure = config.ThrowOnParseFailure;
            int rowNumber = headerRow;

            void Plain(ref TState state, Row row)
            {
                if (row.IsEmptyRecord)
                {
                    return;
                }
                TModel model = info.CreateInstance();
                SparseRowProjection.ParseRow(in row, bindings, default, false, false, provider, throwOnParseFailure, 0, ref model);
                sink(ref state, model);
            }

            void PlainNumbered(ref TState state, Row row)
            {
                int current = ++rowNumber;
                if (row.IsEmptyRecord)
                {
                    return;
                }
                TModel model = info.CreateInstance();
                SparseRowProjection.ParseRow(in row, bindings, default, false, false, provider, throwOnParseFailure, current, ref model);
                sink(ref state, model);
            }

            void Tracked(ref TState state, Row row)
            {
                if (row.IsEmptyRecord)
                {
                    return;
                }
                TModel model = info.CreateInstance();
                if (bindings.Length <= StackallocLimit)
                {
                    Span<bool> seen = stackalloc bool[bindings.Length];
                    SparseRowProjection.ParseRow(in row, bindings, seen, true, false, provider, throwOnParseFailure, 0, ref model);
                }
                else
                {
                    SparseRowProjection.ParseRow(in row, bindings, new bool[bindings.Length], true, false, provider, throwOnParseFailure, 0, ref model);
                }
                sink(ref state, model);
            }

            void TrackedNumbered(ref TState state, Row row)
            {
                int current = ++rowNumber;
                if (row.IsEmptyRecord)
                {
                    return;
                }
                TModel model = info.CreateInstance();
                if (bindings.Length <= StackallocLimit)
                {
                    Span<bool> seen = stackalloc bool[bindings.Length];
                    SparseRowProjection.ParseRow(in row, bindings, seen, true, false, provider, throwOnParseFailure, current, ref model);
                }
                else
                {
                    SparseRowProjection.ParseRow(in row, bindings, new bool[bindings.Length], true, false, provider, throwOnParseFailure, current, ref model);
                }
                sink(ref state, model);
            }

            if (track)
            {
                return sequential ? TrackedNumbered : Tracked;
            }
            return sequential ? PlainNumbered : Plain;
        }
    }
}
