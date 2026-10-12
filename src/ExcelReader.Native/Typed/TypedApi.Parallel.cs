using System.Collections.Concurrent;
using ExcelReader.Core.Parser.ParallelCsv;
using ExcelReader.Core.Parser.SheetPipeline;
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
            if (dop > 1 && sheet >= 0 && sheet < handle.Workbook.SheetCount
                && SheetPipelineSource.TryCreate(handle.Workbook.SheetAt(sheet), dop) is { } pipeline)
            {
                LastParseRanInParallel = true;
                return ParseSheetInParallel(handle, sheet, pipeline, specs, headerRow, dop, chunkSizeOverride, out table);
            }
            return ParseSequential(handle, sheet, specs, headerRow, out table);
        }

        private static int ParseSheetInParallel(NativeHandle handle, int sheet, SheetPipelineSource pipeline, NativeColumnSpec[] specs,
            int headerRow, int dop, int chunkSizeOverride, out NativeTable table)
        {
            table = default;
            if (!TryValidateArguments(specs, headerRow, out string? argumentError))
            {
                NativeApi.SetLastError(argumentError);
                return NativeStatus.InvalidArgument;
            }
            NativeApi.ClearLastError();
            var seeded = new ConcurrentBag<PartitionTable>();
            try
            {
                int[] columnIndices = new int[specs.Length];
                using (IExcelRowEnumerator rows = handle.Workbook.SheetAt(sheet).GetEnumerator())
                {
                    if (!TryResolveColumns(rows, specs, headerRow, columnIndices, out string? resolveError))
                    {
                        NativeApi.SetLastError(resolveError);
                        return NativeStatus.InvalidArgument;
                    }
                }

                bool isDate1904 = handle.Workbook.IsDate1904;
                RowAggregation<PartitionTable> aggregation = new()
                {
                    Seed = () =>
                    {
                        var partition = new PartitionTable(specs);
                        seeded.Add(partition);
                        return partition;
                    },
                    Accumulate = (ref PartitionTable state, Row row) => state.Append(row, columnIndices, isDate1904),
                    Combine = PartitionTable.Combine,
                };
                int chunkBytes = chunkSizeOverride > 0 ? chunkSizeOverride : SheetPipelineSource.DefaultChunkBytes;
                PartitionTable merged = Task.Run(() => SheetPipelineRunner.AggregateAsync(
                    pipeline, aggregation, null, headerRow, dop, chunkBytes, CancellationToken.None)).GetAwaiter().GetResult();
                if (merged.FailedColumn >= 0)
                {
                    NativeApi.SetLastError(DescribeFailedColumn(specs, merged.FailedColumn));
                    return NativeStatus.Error;
                }
                table = BuildTable([.. merged.Parts], dop);
                return NativeStatus.Ok;
            }
            catch (Exception exception)
            {
                NativeApi.SetLastError(exception.Message);
                table = default;
                return NativeStatus.Error;
            }
            finally
            {
                // The runner drops states on failure and recovery, and Combine folds builders into another state:
                // every seeded state is disposed, and disposing a builder twice is a no-op.
                foreach (PartitionTable partition in seeded)
                {
                    partition.Dispose();
                }
            }
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
                RowAggregation<PartitionTable> aggregation = new()
                {
                    Seed = () => new PartitionTable(specs),
                    Accumulate = (ref PartitionTable state, Row row) => state.Append(row, columnIndices, isDate1904),
                    Combine = PartitionTable.Combine,
                };
                PartitionTable merged = Task.Run(() => ParallelCsvProcessor.RunPartitionedAsync(
                    source, aggregation, reader, headerRow, degreeOfParallelism, chunkSizeOverride, CancellationToken.None)).GetAwaiter().GetResult();
                try
                {
                    if (merged.FailedColumn >= 0)
                    {
                        NativeApi.SetLastError(DescribeFailedColumn(specs, merged.FailedColumn));
                        return NativeStatus.Error;
                    }
                    table = BuildTable([.. merged.Parts], ParallelCsvFactory.Normalize(degreeOfParallelism));
                    return NativeStatus.Ok;
                }
                finally
                {
                    merged.Dispose();
                }
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

        private static void DisposeAll(ColumnBuilder[] builders)
        {
            foreach (ColumnBuilder builder in builders)
            {
                builder.Dispose();
            }
        }

        private sealed class PartitionTable : IDisposable
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

            public void Dispose()
            {
                foreach (ColumnBuilder[] part in Parts)
                {
                    DisposeAll(part);
                }
            }

            internal static PartitionTable Combine(PartitionTable left, PartitionTable right)
            {
                if (left.FailedColumn >= 0)
                {
                    right.Dispose();
                    return left;
                }
                if (right.FailedColumn >= 0)
                {
                    left.Dispose();
                    return right;
                }
                left.Parts.AddRange(right.Parts);
                return left;
            }
        }
    }
}
