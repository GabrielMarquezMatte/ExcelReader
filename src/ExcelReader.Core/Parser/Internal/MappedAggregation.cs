using ExcelReader.Core.Reader;
using ExcelReader.Core.ValueObjects;

namespace ExcelReader.Core.Parser.Internal
{
    internal static class MappedAggregation<TAccumulator, TModel>
        where TAccumulator : ICsvAccumulator<TAccumulator, TModel>, new()
        where TModel : allows ref struct
    {
        private const int StackallocLimit = 256;

        internal static readonly CsvAggregation<TAccumulator> Unbound = new()
        {
            Seed = static () => new TAccumulator(),
            Accumulate = static (ref _, _) => throw new InvalidOperationException("The model map was not bound before the first record."),
            Combine = static (left, right) =>
            {
                left.Merge(right);
                return left;
            },
        };

        internal static CsvAccumulateFactory<TAccumulator> Binder(CsvModelMap<TModel> map, int headerRow)
        {
            return (Row header, bool sequential) => Bind(map, header, headerRow, sequential);
        }

        private static CsvRowAction<TAccumulator> Bind(CsvModelMap<TModel> map, Row header, int headerRow, bool sequential)
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

            void Plain(ref TAccumulator accumulator, Row row)
            {
                if (row.IsEmptyRecord)
                {
                    return;
                }
                TModel model = info.CreateInstance();
                SparseRowProjection.ParseRow(in row, bindings, default, false, false, provider, throwOnParseFailure, 0, ref model);
                accumulator.Add(model);
            }

            void PlainNumbered(ref TAccumulator accumulator, Row row)
            {
                int current = ++rowNumber;
                if (row.IsEmptyRecord)
                {
                    return;
                }
                TModel model = info.CreateInstance();
                SparseRowProjection.ParseRow(in row, bindings, default, false, false, provider, throwOnParseFailure, current, ref model);
                accumulator.Add(model);
            }

            void Tracked(ref TAccumulator accumulator, Row row)
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
                accumulator.Add(model);
            }

            void TrackedNumbered(ref TAccumulator accumulator, Row row)
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
                accumulator.Add(model);
            }

            if (track)
            {
                return sequential ? TrackedNumbered : Tracked;
            }
            return sequential ? PlainNumbered : Plain;
        }
    }
}
