using System.Globalization;
using System.Runtime.InteropServices;
using System.Text;
using ExcelReader.Core.Enums;
using ExcelReader.Core.Parser;
using ExcelReader.Core.Reader;
using ExcelReader.Core.ValueObjects;

namespace ExcelReader.Native
{
    internal static unsafe partial class NativeApi
    {
        /// <summary>
        /// Guesses a <see cref="ParseTyped"/>/<c>xl_parse_arrow</c> schema by sampling the WHOLE current
        /// sheet, from its first row — independent of, and never disturbing, the incremental cursor
        /// <see cref="NextRow"/>/<see cref="NextRowDecoded"/>/<see cref="ReadAllBlob"/> share on
        /// <paramref name="handle"/>. Every guess comes from the sampled cells' own <see cref="CellType"/>
        /// tag (the same one <see cref="ParseTyped"/> already trusts to convert values) — no text
        /// sniffing, and no new parsing logic beyond <see cref="ExcelCellReaders.Parsable{TValue}"/>,
        /// reused here only to tell an integral column from a fractional one.
        /// </summary>
        /// <param name="headerRow">Same meaning as in <see cref="ParseTyped"/>: 1-based row number to
        /// take column names from; 0 means "no header", so every returned spec is index-based.</param>
        /// <param name="sampleSize">How many rows after the header to inspect. Must be positive.</param>
        internal static int InferSchema(NativeHandle? handle, int headerRow, int sampleSize, out NativeInferredSchema schema)
        {
            schema = default;
            if (handle is null)
            {
                return NativeStatus.InvalidHandle;
            }
            if (headerRow < 0)
            {
                SetLastError($"header_row must be 0 (no header) or a positive row number; got {headerRow}.");
                return NativeStatus.InvalidArgument;
            }
            if (sampleSize <= 0)
            {
                SetLastError($"sample_size must be positive; got {sampleSize}.");
                return NativeStatus.InvalidArgument;
            }

            ClearLastError();
            IExcelRowEnumerator? rows = null;
            try
            {
                rows = handle.Reader.GetEnumerator();
                List<string?> names = [];
                List<ColumnStat> stats = [];

                if (headerRow > 0 && !TryReadHeader(rows, headerRow, names, stats, out string? headerError))
                {
                    SetLastError(headerError!);
                    return NativeStatus.InvalidArgument;
                }

                bool isDate1904 = handle.Reader.IsDate1904;
                int dataRowCount = SampleDataRows(rows, sampleSize, isDate1904, names, stats);
                MarkSparseColumnsNullable(stats, dataRowCount);

                schema = BuildSchema(names, stats);
                return NativeStatus.Ok;
            }
            catch (Exception exception)
            {
                SetLastError(exception.Message);
                schema = default;
                return NativeStatus.Error;
            }
            finally
            {
                rows?.Dispose();
            }
        }

        /// <summary>Releases a result returned by <see cref="InferSchema"/> and resets it to zero. Safe on a zeroed value.</summary>
        internal static void FreeSchema(ref NativeInferredSchema schema)
        {
            if (schema.Columns == IntPtr.Zero)
            {
                schema = default;
                return;
            }

            NativeColumnSpecRaw* columns = (NativeColumnSpecRaw*)schema.Columns;
            for (int i = 0; i < schema.ColumnCount; i++)
            {
                if (columns[i].Name is not null)
                {
                    Marshal.FreeHGlobal((IntPtr)columns[i].Name);
                }
            }
            Marshal.FreeHGlobal(schema.Columns);
            schema = default;
        }

        // Advances `rows` past headerRow and records each populated header cell's trimmed text as that
        // column's name. A blank header cell leaves its column name null (index-based) rather than an
        // empty string — an empty name would fail xl_parse_typed's own "blank name" validation later.
        private static bool TryReadHeader(IExcelRowEnumerator rows, int headerRow, List<string?> names, List<ColumnStat> stats, out string? error)
        {
            error = null;
            for (int rowNumber = 1; rowNumber <= headerRow; rowNumber++)
            {
                if (!rows.MoveNext())
                {
                    error = $"sheet has fewer than {headerRow} row(s); cannot resolve header_row.";
                    return false;
                }
            }
            Row header = rows.Current;
            foreach (RowCell cell in header.Cells)
            {
                EnsureCapacity(names, stats, cell.ColumnIndex);
                string text = cell.Value.GetString().Trim();
                names[cell.ColumnIndex] = text.Length == 0 ? null : text;
            }
            return true;
        }

        // Reads up to sampleSize rows after the header, folding each populated cell's type into that
        // column's ColumnStat. Returns how many data rows were actually sampled (fewer than sampleSize
        // at end of sheet), used by MarkSparseColumnsNullable below.
        private static int SampleDataRows(IExcelRowEnumerator rows, int sampleSize, bool isDate1904, List<string?> names, List<ColumnStat> stats)
        {
            int dataRowCount = 0;
            for (int i = 0; i < sampleSize && rows.MoveNext(); i++)
            {
                dataRowCount++;
                Row row = rows.Current;
                foreach (RowCell cell in row.Cells)
                {
                    EnsureCapacity(names, stats, cell.ColumnIndex);
                    Cell value = cell.Value;
                    CollectionsMarshal.AsSpan(stats)[cell.ColumnIndex].Observe(in value, isDate1904);
                }
            }
            return dataRowCount;
        }

        // A column absent from some sampled rows never triggers ColumnStat.Observe for them — which is
        // exactly what row[index] would have reported for it: CellType.Empty. Comparing each column's
        // Observe count against the number of rows actually sampled catches that without a second,
        // O(columns) walk of every row via the indexer.
        private static void MarkSparseColumnsNullable(List<ColumnStat> stats, int dataRowCount)
        {
            foreach (ref ColumnStat stat in CollectionsMarshal.AsSpan(stats))
            {
                if (stat.SeenCount < dataRowCount)
                {
                    stat.SawEmpty = true;
                }
            }
        }

        private static void EnsureCapacity(List<string?> names, List<ColumnStat> stats, int index)
        {
            while (names.Count <= index)
            {
                names.Add(null);
                stats.Add(default);
            }
        }

        // Every allocation this function makes is handed to the caller inside the returned schema and
        // freed only by FreeSchema — a thrown exception between an AllocHGlobal and the assignment to
        // `schema` below would leak it, but nothing after the loop's own allocations can throw.
        private static NativeInferredSchema BuildSchema(List<string?> names, List<ColumnStat> stats)
        {
            int columnCount = names.Count;
            if (columnCount == 0)
            {
                return new NativeInferredSchema { Columns = IntPtr.Zero, ColumnCount = 0 };
            }

            NativeColumnSpecRaw* block = (NativeColumnSpecRaw*)Marshal.AllocHGlobal(checked(columnCount * sizeof(NativeColumnSpecRaw)));
            for (int i = 0; i < columnCount; i++)
            {
                block[i] = BuildSpec(names[i], i, stats[i]);
            }
            return new NativeInferredSchema { Columns = (IntPtr)block, ColumnCount = columnCount };
        }

        private static NativeColumnSpecRaw BuildSpec(string? name, int index, ColumnStat stat)
        {
            byte* namePtr = null;
            int nameLen = 0;
            if (name is not null)
            {
                nameLen = Encoding.UTF8.GetByteCount(name);
                namePtr = (byte*)Marshal.AllocHGlobal(Math.Max(nameLen, 1));
                Encoding.UTF8.GetBytes(name, new Span<byte>(namePtr, nameLen));
            }
            return new NativeColumnSpecRaw
            {
                Name = namePtr,
                NameLen = nameLen,
                Index = index,
                Type = stat.InferType(),
                Nullable = stat.SawEmpty ? 1 : 0,
            };
        }

        // Accumulates what kinds of CellType a column's sampled cells held, enough to guess a
        // xl_parse_typed ColumnSpec without ever storing a cell's value past its own row.
        private struct ColumnStat
        {
            internal int SeenCount;
            internal bool SawEmpty;
            internal bool SawString;
            internal bool SawNumber;
            internal bool SawDate;
            internal bool SawBool;
            internal bool SawFormulaOrError;
            internal bool SawNonIntegralNumber;

            internal void Observe(in Cell cell, bool isDate1904)
            {
                SeenCount++;
                switch (cell.Type)
                {
                    case CellType.Empty:
                        SawEmpty = true;
                        break;
                    case CellType.ExcelString:
                        SawString = true;
                        break;
                    case CellType.Number:
                        SawNumber = true;
                        if (!ExcelCellReaders.Parsable<long>(in cell, isDate1904, CultureInfo.InvariantCulture, out _))
                        {
                            SawNonIntegralNumber = true;
                        }
                        break;
                    case CellType.Date:
                        SawDate = true;
                        break;
                    case CellType.Boolean:
                        SawBool = true;
                        break;
                    default: // Formula, Error — the cached result was never sampled as a plain value.
                        SawFormulaOrError = true;
                        break;
                }
            }

            // A column only gets a non-string guess when every sampled cell agreed on one single kind.
            // A real mix, any formula/error result, or nothing seen at all falls back to the string
            // type, since it is the only one able to represent every one of those verbatim.
            internal readonly int InferType()
            {
                int kinds = (SawString ? 1 : 0) + (SawNumber ? 1 : 0) + (SawDate ? 1 : 0) + (SawBool ? 1 : 0);
                if (SawFormulaOrError || kinds != 1)
                {
                    return NativeColumnType.String;
                }
                if (SawString)
                {
                    return NativeColumnType.String;
                }
                if (SawDate)
                {
                    return NativeColumnType.Date;
                }
                if (SawBool)
                {
                    return NativeColumnType.Bool;
                }
                return SawNonIntegralNumber ? NativeColumnType.Float64 : NativeColumnType.Int64;
            }
        }
    }

    /// <summary>Flat C ABI representation of the whole result of <see cref="NativeApi.InferSchema"/>.
    /// <see cref="Columns"/> is one allocation of <see cref="ColumnCount"/> <see cref="NativeColumnSpecRaw"/>
    /// values; each spec's own non-null <see cref="NativeColumnSpecRaw.Name"/> is a separate allocation,
    /// freed individually by <see cref="NativeApi.FreeSchema"/>.</summary>
    [StructLayout(LayoutKind.Sequential)]
    internal struct NativeInferredSchema
    {
        public IntPtr Columns;
        public int ColumnCount;
    }
}
