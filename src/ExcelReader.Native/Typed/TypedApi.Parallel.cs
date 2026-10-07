using ExcelReader.Core.Parser.ParallelCsv;
using ExcelReader.Core.Reader;
using ExcelReader.Core.Reader.Csv;

namespace ExcelReader.Native.Typed
{
    internal static unsafe partial class TypedApi
    {
        [ThreadStatic]
        internal static bool LastParseRanInParallel;

        internal static int ParseTypedTable(NativeHandle? handle, int sheet, NativeColumnSpec[] specs, int headerRow, int degreeOfParallelism,
            out NativeTable table, int chunkSizeOverride = 0)
        {
            table = default;
            LastParseRanInParallel = false;
            if (handle is null)
            {
                return NativeStatus.InvalidHandle;
            }
            if (degreeOfParallelism < 0)
            {
                NativeApi.SetLastError($"degree_of_parallelism must be 0 (processor count) or positive; got {degreeOfParallelism}.");
                return NativeStatus.InvalidArgument;
            }
            int dop = ParallelCsvFactory.Normalize(degreeOfParallelism);
            if (sheet == 0 && dop > 1 && handle.Workbook is CsvReader csv && csv.TryGetChunkSource(out CsvChunkSource source)
                && ParallelCsvFactory.CanPartition(dop, source.Length, csv.Options))
            {
                LastParseRanInParallel = true;
                return ParseCsvInParallel(handle, source, csv.Options, specs, headerRow, degreeOfParallelism, chunkSizeOverride, out table);
            }
            return ParseSequential(handle, sheet, specs, headerRow, out table);
        }

        private static int ParseCsvInParallel(NativeHandle handle, CsvChunkSource source, CsvReaderOptions reader, NativeColumnSpec[] specs,
            int headerRow, int degreeOfParallelism, int chunkSizeOverride, out NativeTable table)
        {
            table = default;
            if (!TryValidateArguments(specs, headerRow, out string? argumentError))
            {
                NativeApi.SetLastError(argumentError);
                return NativeStatus.InvalidArgument;
            }
            NativeApi.ClearLastError();
            try
            {
                int[] columnIndices = new int[specs.Length];
                using (CsvReader headerReader = source.OpenReader(reader))
                using (IExcelRowEnumerator rows = headerReader.OpenSheet())
                {
                    if (!TryResolveColumns(rows, specs, headerRow, columnIndices, out string? resolveError))
                    {
                        NativeApi.SetLastError(resolveError);
                        return NativeStatus.InvalidArgument;
                    }
                }

                bool isDate1904 = handle.Workbook.IsDate1904;
                CsvAggregation<PartitionTable> aggregation = new()
                {
                    Seed = () => new PartitionTable(specs),
                    Accumulate = (ref PartitionTable state, Row row) => state.Append(row, columnIndices, isDate1904),
                    Combine = PartitionTable.Combine,
                };
                CsvParallelOptions options = new() { DegreeOfParallelism = degreeOfParallelism, HeaderRow = headerRow, Reader = reader };
                PartitionTable merged = Task.Run(() => ParallelCsvProcessor.RunPartitionedAsync(
                    source, aggregation, options, chunkSizeOverride, CancellationToken.None)).GetAwaiter().GetResult();
                if (merged.FailedColumn >= 0)
                {
                    NativeApi.SetLastError(DescribeFailedColumn(specs, merged.FailedColumn));
                    return NativeStatus.Error;
                }
                table = BuildTable([.. merged.Parts], ParallelCsvFactory.Normalize(degreeOfParallelism));
                return NativeStatus.Ok;
            }
            catch (Exception exception)
            {
                NativeApi.SetLastError(exception.Message);
                table = default;
                return NativeStatus.Error;
            }
        }

        private static ColumnBuilder[] NewBuilders(NativeColumnSpec[] specs)
        {
            ColumnBuilder[] builders = new ColumnBuilder[specs.Length];
            for (int i = 0; i < specs.Length; i++)
            {
                builders[i] = new ColumnBuilder(specs[i].Type, specs[i].Nullable);
            }
            return builders;
        }

        private sealed class PartitionTable
        {
            private readonly ColumnBuilder[] _builders;

            internal PartitionTable(NativeColumnSpec[] specs)
            {
                _builders = NewBuilders(specs);
                Parts = [_builders];
            }

            internal List<ColumnBuilder[]> Parts { get; }

            internal int FailedColumn { get; private set; } = -1;

            internal void Append(in Row row, int[] columnIndices, bool isDate1904)
            {
                if (FailedColumn < 0 && !TryAppendRow(_builders, row, columnIndices, isDate1904, out int failed))
                {
                    FailedColumn = failed;
                }
            }

            internal static PartitionTable Combine(PartitionTable left, PartitionTable right)
            {
                if (left.FailedColumn >= 0)
                {
                    return left;
                }
                if (right.FailedColumn >= 0)
                {
                    return right;
                }
                left.Parts.AddRange(right.Parts);
                return left;
            }
        }
    }
}
