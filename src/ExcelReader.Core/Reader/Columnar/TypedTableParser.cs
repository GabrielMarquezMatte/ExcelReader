using System.Collections.Concurrent;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using ExcelReader.Core.Parser.ParallelCsv;
using ExcelReader.Core.Parser.SheetPipeline;
using ExcelReader.Core.Reader.Csv;
using ExcelReader.Core.Reader.Schema;

namespace ExcelReader.Core.Reader.Columnar
{
    /// <summary>
    /// Reads a sheet into typed columns, on one thread or across partitions. Shared by the C ABI's
    /// typed and Arrow exports and by ExcelReader.Arrow, so they convert every cell the same way.
    /// </summary>
    internal static class TypedTableParser
    {
        /// <summary>
        /// Parses <paramref name="sheet"/>. A spec that does not fit the sheet returns
        /// <see langword="null"/> with <paramref name="specError"/>; a value that fails to convert on
        /// a non-nullable column throws <see cref="InvalidOperationException"/>.
        /// </summary>
        internal static ColumnarTable? Parse(IExcelSheet sheet, TypedColumnSpec[] specs, int headerRow, int degreeOfParallelism,
            bool allowDictionary, int chunkSizeOverride, out bool ranInParallel, out string? specError)
        {
            ranInParallel = false;
            if (!TryValidate(specs, headerRow, allowDictionary, out specError))
            {
                return null;
            }
            int dop = ParallelCsvFactory.Normalize(degreeOfParallelism);
            if (dop > 1 && sheet is CsvSheet csv && csv.Reader.TryGetChunkSource(out CsvChunkSource source)
                && ParallelCsvFactory.CanPartition(dop, source.Length, csv.Reader.Options))
            {
                ranInParallel = true;
                return ParseCsvInParallel(source, csv.Reader.Options, sheet.IsDate1904, specs, headerRow, dop, chunkSizeOverride, out specError);
            }
            if (dop > 1 && SheetPipelineSource.TryCreate(sheet, dop) is { } pipeline)
            {
                ranInParallel = true;
                return ParseSheetInParallel(sheet, pipeline, specs, headerRow, dop, chunkSizeOverride, out specError);
            }
            return ParseSequential(sheet, specs, headerRow, out specError);
        }

        private static ColumnarTable? ParseSequential(IExcelSheet sheet, TypedColumnSpec[] specs, int headerRow, out string? specError)
        {
            using IExcelRowEnumerator rows = sheet.GetEnumerator();
            int[] columnIndices = new int[specs.Length];
            if (!TryResolveColumns(rows, specs, headerRow, columnIndices, out specError))
            {
                return null;
            }
            using TypedTableBuilder builder = new(specs);
            bool isDate1904 = sheet.IsDate1904;
            while (rows.MoveNext())
            {
                if (!builder.TryAppend(rows.Current, columnIndices, isDate1904, out int failedColumn))
                {
                    throw new InvalidOperationException(DescribeFailedColumn(specs, failedColumn));
                }
            }
            return builder.Build();
        }

        private static ColumnarTable? ParseSheetInParallel(IExcelSheet sheet, SheetPipelineSource pipeline, TypedColumnSpec[] specs,
            int headerRow, int dop, int chunkSizeOverride, out string? specError)
        {
            int[] columnIndices = new int[specs.Length];
            using (IExcelRowEnumerator rows = sheet.GetEnumerator())
            {
                if (!TryResolveColumns(rows, specs, headerRow, columnIndices, out specError))
                {
                    return null;
                }
            }
            ConcurrentBag<PartitionTable> seeded = [];
            try
            {
                RowAggregation<PartitionTable> aggregation = Aggregation(specs, columnIndices, sheet.IsDate1904, seeded);
                int chunkBytes = chunkSizeOverride > 0 ? chunkSizeOverride : SheetPipelineSource.DefaultChunkBytes;
                PartitionTable merged = Task.Run(() => SheetPipelineRunner.AggregateAsync(
                    pipeline, aggregation, null, headerRow, dop, chunkBytes, CancellationToken.None)).GetAwaiter().GetResult();
                return Finish(merged, specs, dop);
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

        private static ColumnarTable? ParseCsvInParallel(CsvChunkSource source, CsvReaderOptions reader, bool isDate1904, TypedColumnSpec[] specs,
            int headerRow, int dop, int chunkSizeOverride, out string? specError)
        {
            int[] columnIndices = new int[specs.Length];
            using (CsvReader headerReader = source.OpenReader(reader))
            using (IExcelRowEnumerator rows = headerReader.OpenSheet())
            {
                if (!TryResolveColumns(rows, specs, headerRow, columnIndices, out specError))
                {
                    return null;
                }
            }
            ConcurrentBag<PartitionTable> seeded = [];
            try
            {
                RowAggregation<PartitionTable> aggregation = Aggregation(specs, columnIndices, isDate1904, seeded);
                PartitionTable merged = Task.Run(() => ParallelCsvProcessor.RunPartitionedAsync(
                    source, aggregation, reader, headerRow, dop, chunkSizeOverride, CancellationToken.None)).GetAwaiter().GetResult();
                return Finish(merged, specs, dop);
            }
            finally
            {
                foreach (PartitionTable partition in seeded)
                {
                    partition.Dispose();
                }
            }
        }

        private static RowAggregation<PartitionTable> Aggregation(TypedColumnSpec[] specs, int[] columnIndices, bool isDate1904,
            ConcurrentBag<PartitionTable> seeded)
        {
            return new RowAggregation<PartitionTable>
            {
                Seed = () =>
                {
                    PartitionTable partition = new(specs);
                    seeded.Add(partition);
                    return partition;
                },
                Accumulate = (ref PartitionTable state, Row row) => state.Append(row, columnIndices, isDate1904),
                Combine = PartitionTable.Combine,
            };
        }

        private static ColumnarTable Finish(PartitionTable merged, TypedColumnSpec[] specs, int dop)
        {
            if (merged.FailedColumn >= 0)
            {
                throw new InvalidOperationException(DescribeFailedColumn(specs, merged.FailedColumn));
            }
            return TypedTableBuilder.Build(merged.Parts, dop);
        }

        internal static string DescribeFailedColumn(TypedColumnSpec[] specs, int failedColumn)
        {
            TypedColumnSpec spec = specs[failedColumn];
            string columnLabel = spec.Names.Length > 0
                ? string.Join(" / ", spec.Names)
                : spec.Label ?? spec.Index.ToString(CultureInfo.InvariantCulture);
            return $"column {failedColumn} (\"{columnLabel}\") has a value that failed to convert and is not nullable.";
        }

        internal static bool TryValidate(TypedColumnSpec[] specs, int headerRow, bool allowDictionary, [NotNullWhen(false)] out string? error)
        {
            error = null;
            if (headerRow < 0)
            {
                error = $"header_row must be 0 (no header) or a positive row number; got {headerRow}.";
                return false;
            }
            foreach (TypedColumnSpec spec in specs)
            {
                error = SpecError(spec, headerRow, allowDictionary);
                if (error is not null)
                {
                    return false;
                }
            }
            return true;
        }

        private static string? SpecError(TypedColumnSpec spec, int headerRow, bool allowDictionary)
        {
            if (spec.Names.Length == 0 && spec.Index < 0)
            {
                return "a column spec with no name must have a non-negative index.";
            }
            if (spec.Names.Length > 0 && headerRow == 0)
            {
                return $"column \"{spec.Names[0]}\" is name-based, but header_row is 0 (no header row to match it against).";
            }
            foreach (string name in spec.Names)
            {
                if (name.AsSpan().Trim().IsEmpty)
                {
                    return "a name-based column spec cannot have a blank name.";
                }
            }
            if (spec.Type is < ExcelColumnType.StringColumn or > ExcelColumnType.StringDictionaryColumn)
            {
                return $"column spec has unknown type {(int)spec.Type}.";
            }
            if (spec.Type == ExcelColumnType.StringDictionaryColumn && !allowDictionary)
            {
                return "a dictionary-encoded string column is only supported by the Arrow outputs.";
            }
            return null;
        }

        internal static bool TryResolveColumns(IExcelRowEnumerator rows, TypedColumnSpec[] specs, int headerRow, int[] columnIndices,
            [NotNullWhen(false)] out string? error)
        {
            error = null;
            if (headerRow == 0)
            {
                for (int i = 0; i < specs.Length; i++)
                {
                    columnIndices[i] = specs[i].Index;
                }
                return true;
            }
            if (!SchemaInference.TrySkipToHeaderRow(rows, headerRow, out error))
            {
                return false;
            }

            Dictionary<string, int>.AlternateLookup<ReadOnlySpan<char>>? header = null;
            for (int i = 0; i < specs.Length; i++)
            {
                string[] names = specs[i].Names;
                if (names.Length == 0)
                {
                    columnIndices[i] = specs[i].Index;
                    continue;
                }
                header ??= HeaderColumns(rows.Current).GetAlternateLookup<ReadOnlySpan<char>>();
                int found = FindHeader(header.Value, names);
                if (found < 0)
                {
                    error = $"no column header matches any of {string.Join(", ", Array.ConvertAll(names, n => $"\"{n}\""))}.";
                    return false;
                }
                columnIndices[i] = found;
            }
            return true;
        }

        private static int FindHeader(Dictionary<string, int>.AlternateLookup<ReadOnlySpan<char>> header, string[] names)
        {
            foreach (string name in names)
            {
                if (header.TryGetValue(name.AsSpan().Trim(), out int found))
                {
                    return found;
                }
            }
            return -1;
        }

        // The first column carrying a header wins, as a left-to-right search would find it.
        private static Dictionary<string, int> HeaderColumns(Row header)
        {
            Dictionary<string, int> columns = new(StringComparer.OrdinalIgnoreCase);
            foreach (RowCell cell in header.Cells)
            {
                columns.TryAdd(cell.Value.GetString().Trim(), cell.ColumnIndex);
            }
            return columns;
        }

        private sealed class PartitionTable : IDisposable
        {
            internal PartitionTable(TypedColumnSpec[] specs)
            {
                Parts = [new TypedTableBuilder(specs)];
            }

            internal List<TypedTableBuilder> Parts { get; }

            internal int FailedColumn { get; private set; } = -1;

            internal void Append(in Row row, int[] columnIndices, bool isDate1904)
            {
                if (FailedColumn < 0 && !Parts[0].TryAppend(row, columnIndices, isDate1904, out int failed))
                {
                    FailedColumn = failed;
                }
            }

            public void Dispose()
            {
                foreach (TypedTableBuilder part in Parts)
                {
                    part.Dispose();
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
