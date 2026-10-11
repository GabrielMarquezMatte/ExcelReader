using System.Runtime.CompilerServices;
using ExcelReader.Core.Parser.Internal;
using ExcelReader.Core.Parser.ParallelCsv;
using ExcelReader.Core.Reader;

namespace ExcelReader.Core.Parser.SheetPipeline
{
    internal static class SheetRowProjection<T>
    {
        internal sealed class Chunk
        {
            public readonly List<T> Items = [];
            public int Rows;
            public bool[]? Seen;
        }

        private static readonly RowAggregation<Chunk> Unbound = new()
        {
            Seed = static () => new Chunk(),
            Accumulate = static (ref _, _) => throw new InvalidOperationException("The model map was not bound before the first row."),
            Combine = static (left, _) => left,
        };

        internal static async IAsyncEnumerable<T> ParseAsync(
            SheetPipelineSource source, ExcelParser<T> parser, bool isDate1904, int dop, int chunkBytes, [EnumeratorCancellation] CancellationToken ct)
        {
            TypeMapInfo<T> info = parser.Info;
            int headerRow = info.IsIndexBased ? 0 : parser.Config.HeaderRow;
            await foreach (Chunk chunk in SheetPipelineRunner.RunAsync(source, Unbound, Binder(parser, info, isDate1904), headerRow, dop, chunkBytes, ct).ConfigureAwait(false))
            {
                foreach (T item in chunk.Items)
                {
                    yield return item;
                }
            }
        }

        private static RowActionFactory<Chunk> Binder(ExcelParser<T> parser, TypeMapInfo<T> info, bool isDate1904)
        {
            return (Row header, bool _) =>
            {
                ExcelParserOptions config = parser.Config;
                ColumnBinding<T>[] bindings = info.IsIndexBased
                    ? info.IndexBindings
                    : SparseRowProjection.BuildColumnMap(in header, info, config.ColumnNameComparer, config.HeaderNormalization, out int _);
                bool track = bindings.Any(static binding => binding.RequireValue);
                ExcelRowContext context = new() { IsDate1904 = isDate1904, FormatProvider = config.Culture };
                bool throwOnParseFailure = config.ThrowOnParseFailure;
                return (ref Chunk chunk, Row row) =>
                {
                    int rowInChunk = ++chunk.Rows;
                    T model = info.CreateInstance();
                    Span<bool> seen = track ? (chunk.Seen ??= new bool[bindings.Length]) : default;
                    SparseRowProjection.ParseRow(in row, bindings, seen, track, in context, throwOnParseFailure, rowInChunk, ref model);
                    chunk.Items.Add(model);
                };
            };
        }
    }
}
