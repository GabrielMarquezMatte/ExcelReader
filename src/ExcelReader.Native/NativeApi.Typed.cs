using System.Globalization;
using System.Runtime.InteropServices;
using ExcelReader.Core.Parser;
using ExcelReader.Core.Reader;
using ExcelReader.Core.ValueObjects;

namespace ExcelReader.Native
{
    internal static unsafe partial class NativeApi
    {
        /// <summary>
        /// Schema-driven columnar read of the WHOLE current sheet, from its first row — independent of,
        /// and never disturbing, the incremental cursor <see cref="NextRow"/>/<see cref="NextRowDecoded"/>/
        /// <see cref="ReadAllBlob"/> share on <paramref name="handle"/>. This is not a new parser: each
        /// column dispatches to the same <see cref="ExcelCellReaders"/> members
        /// <c>ExcelParser&lt;T&gt;</c>'s reflective path already uses (see
        /// docs/NATIVE_BINDINGS_PLAN.md §7's feasibility finding).
        /// </summary>
        /// <param name="headerRow">1-based row number to resolve name-based <paramref name="specs"/>
        /// against; rows before it are skipped entirely and it is never itself yielded as data. 0 means
        /// "no header" — every row from the first is data, and every spec must be index-based.</param>
        internal static int ParseTyped(NativeHandle? handle, NativeColumnSpec[] specs, int headerRow, out NativeTable table)
        {
            table = default;
            if (handle is null)
            {
                return NativeStatus.InvalidHandle;
            }
            if (!TryValidateArguments(specs, headerRow, out string? argumentError))
            {
                SetLastError(argumentError!);
                return NativeStatus.InvalidArgument;
            }

            ClearLastError();
            IExcelRowEnumerator? rows = null;
            try
            {
                rows = handle.Reader.GetEnumerator();
                int[] columnIndices = new int[specs.Length];
                if (!TryResolveColumns(rows, specs, headerRow, columnIndices, out string? resolveError))
                {
                    SetLastError(resolveError!);
                    return NativeStatus.InvalidArgument;
                }

                ColumnBuilder[] builders = new ColumnBuilder[specs.Length];
                for (int i = 0; i < specs.Length; i++)
                {
                    builders[i] = new ColumnBuilder(specs[i].Type, specs[i].Nullable);
                }

                bool isDate1904 = handle.Reader.IsDate1904;
                while (rows.MoveNext())
                {
                    if (!TryAppendRow(builders, rows.Current, columnIndices, isDate1904, out int failedColumn))
                    {
                        NativeColumnSpec spec = specs[failedColumn];
                        string columnLabel = spec.Names.Length > 0 ? string.Join(" / ", spec.Names) : spec.Index.ToString(CultureInfo.InvariantCulture);
                        SetLastError($"column {failedColumn} (\"{columnLabel}\") has a value that failed to convert and is not nullable.");
                        return NativeStatus.Error;
                    }
                }

                table = BuildTable(builders);
                return NativeStatus.Ok;
            }
            catch (Exception exception)
            {
                SetLastError(exception.Message);
                table = default;
                return NativeStatus.Error;
            }
            finally
            {
                rows?.Dispose();
            }
        }

        // Split out of ParseTyped's row loop to keep that loop inside the style guide's three-level
        // nesting limit. The failing column travels back through `failedColumn` rather than being
        // reported here, because only the caller holds the specs needed to name it in the message.
        private static bool TryAppendRow(ColumnBuilder[] builders, in Row row, int[] columnIndices, bool isDate1904, out int failedColumn)
        {
            for (int i = 0; i < builders.Length; i++)
            {
                if (!builders[i].AppendFrom(row[columnIndices[i]], isDate1904))
                {
                    failedColumn = i;
                    return false;
                }
            }
            failedColumn = -1;
            return true;
        }

        /// <summary>Releases a result returned by <see cref="ParseTyped"/> and resets it to zero. Safe on a zeroed value.</summary>
        internal static void FreeTable(ref NativeTable table)
        {
            if (table.Columns == IntPtr.Zero)
            {
                table = default;
                return;
            }

            for (int index = 0; index < table.ColumnCount; index++)
            {
                NativeColumn column = ColumnAt(table, index);
                // Data is an interior pointer into Values for string columns (see NativeColumn's doc
                // comment) - freeing it here would be a double free, so only Values and Validity are
                // ever independently allocated.
                if (column.Values != IntPtr.Zero)
                {
                    Marshal.FreeHGlobal(column.Values);
                }
                if (column.Validity != IntPtr.Zero)
                {
                    Marshal.FreeHGlobal(column.Validity);
                }
            }
            Marshal.FreeHGlobal(table.Columns);
            table = default;
        }

        /// <summary>
        /// Bounds the spec count xl_parse_typed/xl_parse_arrow receive before it sizes an array and
        /// drives a walk over the caller's spec block.
        /// </summary>
        /// <remarks>
        /// Internal rather than private so tests can pin the boundary directly: the
        /// [UnmanagedCallersOnly] entry points that enforce this cannot be invoked from managed code,
        /// so the predicate is the only part of that guard a unit test can reach. The C smoke test
        /// covers the entry points themselves.
        /// </remarks>
        internal static bool IsValidSpecCount(int specCount)
        {
            return specCount is > 0 and <= NativeLimits.MaxColumnSpecs;
        }

        /// <summary>Bounds one spec's name length before it becomes a read length over caller memory.</summary>
        internal static bool IsValidNameLength(int nameLength)
        {
            return nameLength is >= 0 and <= NativeLimits.MaxColumnNameBytes;
        }

        /// <summary>Bounds one spec's candidate-name count before it sizes an array and drives a walk over the caller's spec block.</summary>
        internal static bool IsValidNameCount(int nameCount)
        {
            return nameCount is >= 0 and <= NativeLimits.MaxNamesPerSpec;
        }

        private static bool TryValidateArguments(NativeColumnSpec[] specs, int headerRow, out string? error)
        {
            error = null;
            if (specs.Length == 0)
            {
                error = "xl_parse_typed requires at least one column spec.";
                return false;
            }
            if (headerRow < 0)
            {
                error = $"header_row must be 0 (no header) or a positive row number; got {headerRow}.";
                return false;
            }
            foreach (NativeColumnSpec spec in specs)
            {
                if (spec.Names.Length == 0 && spec.Index < 0)
                {
                    error = "a column spec with no name must have a non-negative index.";
                    return false;
                }
                if (spec.Names.Length > 0 && headerRow == 0)
                {
                    error = $"column \"{spec.Names[0]}\" is name-based, but header_row is 0 (no header row to match it against).";
                    return false;
                }
                // FindHeaderColumn trims before comparing, so a blank name would match the first empty
                // header cell — resolving to a column the caller never asked for instead of failing.
                foreach (string name in spec.Names)
                {
                    if (name.AsSpan().Trim().IsEmpty)
                    {
                        error = "a name-based column spec cannot have a blank name.";
                        return false;
                    }
                }
                if (spec.Type is < NativeColumnType.String or > NativeColumnType.Timestamp)
                {
                    error = $"column spec has unknown type {spec.Type}.";
                    return false;
                }
            }
            return true;
        }

        // Advances `rows` so that `rows.Current` is the header row itself. Shared with
        // NativeApi.Schema.cs's TryReadHeader so the two cannot drift on the row arithmetic or on the
        // message a too-short sheet produces.
        private static bool TrySkipToHeaderRow(IExcelRowEnumerator rows, int headerRow, out string? error)
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
            return true;
        }

        // Advances `rows` past any skipped rows and the header row itself (headerRow > 0), or leaves it
        // untouched at the sheet's first row (headerRow == 0, index-only specs). Either way, `rows` is
        // positioned so the next MoveNext() yields the first DATA row.
        private static bool TryResolveColumns(IExcelRowEnumerator rows, NativeColumnSpec[] specs, int headerRow, int[] columnIndices, out string? error)
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

            if (!TrySkipToHeaderRow(rows, headerRow, out error))
            {
                return false;
            }

            Row header = rows.Current;
            for (int i = 0; i < specs.Length; i++)
            {
                string[] names = specs[i].Names;
                if (names.Length == 0)
                {
                    columnIndices[i] = specs[i].Index;
                    continue;
                }
                int found = -1;
                foreach (string name in names)
                {
                    found = FindHeaderColumn(header, name);
                    if (found >= 0)
                    {
                        break;
                    }
                }
                if (found < 0)
                {
                    error = $"no column header matches any of {FormatCandidates(names)}.";
                    return false;
                }
                columnIndices[i] = found;
            }
            return true;
        }

        private static string FormatCandidates(string[] names)
        {
            return string.Join(", ", Array.ConvertAll(names, n => $"\"{n}\""));
        }

        // Mirrors ExcelParserConfig's own defaults (HeaderNormalization.Trim, matched with
        // StringComparer.OrdinalIgnoreCase), reimplemented with public APIs only:
        // HeaderNormalizationExtensions.Apply is internal to ExcelReader.Core, and this project has no
        // InternalsVisibleTo access to it. Only Trim is mirrored — a caller who configures
        // CollapseSpaces or RemoveDiacritics on the managed side gets no equivalent here.
        private static int FindHeaderColumn(Row header, string name)
        {
            string target = name.Trim();
            foreach (RowCell cell in header.Cells)
            {
                if (string.Equals(cell.Value.GetString().Trim(), target, StringComparison.OrdinalIgnoreCase))
                {
                    return cell.ColumnIndex;
                }
            }
            return -1;
        }

        private static NativeTable BuildTable(ColumnBuilder[] builders)
        {
            int columnCount = builders.Length;
            long rowCount = columnCount > 0 ? builders[0].RowCount : 0;
            IntPtr columnsBlock = Marshal.AllocHGlobal(checked(columnCount * sizeof(NativeColumn)));
            NativeColumn* columns = (NativeColumn*)columnsBlock;
            for (int i = 0; i < columnCount; i++)
            {
                columns[i] = builders[i].Build();
            }
            return new NativeTable { ColumnCount = columnCount, RowCount = rowCount, Columns = columnsBlock };
        }

        // Every buffer this layer hands across the boundary is the same move: one native block holding a
        // verbatim copy of an accumulated column. Never returns a zero-size allocation, since a column
        // with no rows still needs a non-null pointer the caller can free.
        private static IntPtr CopyToNativeBlock<T>(ChunkedBuffer<T> source) where T : unmanaged
        {
            int byteLength = source.ByteLength;
            IntPtr block = Marshal.AllocHGlobal(Math.Max(byteLength, 1));
            source.CopyTo(new Span<byte>((void*)block, byteLength));
            return block;
        }

        // Arrow's canonical boolean layout is one LSB-first bit per row — the same shape a validity
        // bitmap has. ColumnBuilder accumulates validity already packed, so this is only for converting
        // a one-byte-per-row bool column on the way out to Arrow.
        private static IntPtr PackBitsLsbFirst(ReadOnlySpan<byte> flags)
        {
            int byteLength = Math.Max((flags.Length + 7) / 8, 1);
            IntPtr block = Marshal.AllocHGlobal(byteLength);
            Span<byte> packed = new((void*)block, byteLength);
            packed.Clear();
            for (int i = 0; i < flags.Length; i++)
            {
                if (flags[i] != 0)
                {
                    packed[i >> 3] |= (byte)(1 << (i & 7));
                }
            }
            return block;
        }

        /// <summary>
        /// Accumulates one column's values in managed memory as rows are read, then marshals to a single
        /// <see cref="NativeColumn"/> in <see cref="Build"/> once every row has been read successfully —
        /// deferring native allocation until success is certain means a conversion failure mid-sheet
        /// (<see cref="AppendFrom"/> returning <see langword="false"/>) never has to unwind any native
        /// memory, unlike <see cref="ReadAllDecoded"/>'s per-row native allocations.
        /// </summary>
        private sealed class ColumnBuilder(int type, bool nullable)
        {
            // Already in the layout the ABI hands out: one LSB-first bit per row, 1 = valid, 0 = null.
            // Accumulating packed rather than a byte per row keeps this eight times smaller for a tall
            // sheet, so Build just copies it instead of converting. Together with accumulating Date as
            // int rather than long, this took NativeTypedParseBenchmark from 23.61 MB down to 20.58 MB
            // allocated; wall clock did not move, since neither change removes work, only bytes.
            private readonly ChunkedBuffer<byte> _validity = new();
            private int _rowCount;
            private bool _anyNull;

            // Only one of these is populated, chosen by `type` — see AppendFrom. ChunkedBuffer rather
            // than List<T> because a List regrows by copying the whole column and discarding the old
            // array; see ChunkedBuffer's remarks. Measured by NativeTypedParseBenchmark on
            // Data/65K_Records_Data.xlsb (14 columns, Ryzen 7 5700X, .NET 10.0.10, --job Medium):
            // 20.58 MB -> 11.73 MB allocated and 48.58 ms -> 43.96 ms. The allocation figure is the
            // load-bearing one; the timing gap is smaller than the run-to-run error on either side.
            private readonly ChunkedBuffer<long> _longs = new(); // Int64, Time, Timestamp — all 8-byte
            private readonly ChunkedBuffer<int> _ints = new(); // Date, which the ABI defines as a 4-byte day count
            private readonly ChunkedBuffer<double> _doubles = new(); // Float64
            private readonly ChunkedBuffer<byte> _bools = new(); // Bool, one byte (0/1) per row
            private readonly ChunkedBuffer<int> _stringOffsets = NewStringOffsets(type); // String
            private readonly ChunkedBuffer<byte> _stringData = new(); // String

            // A string column's offset array is one longer than its row count, opening with 0. Seeded
            // here rather than in the field initializer so a non-string column never allocates the
            // buffer's first chunk for an entry it will never use.
            private static ChunkedBuffer<int> NewStringOffsets(int type)
            {
                ChunkedBuffer<int> offsets = new();
                if (type == NativeColumnType.String)
                {
                    offsets.Add(0);
                }
                return offsets;
            }

            // Reused across every row of a string column: one cell's UTF-8 lands here before being
            // appended, so no per-row buffer is allocated. Grows to the widest cell seen, never shrinks.
            private byte[] _scratch = [];

            internal int RowCount
            {
                get
                {
                    return _rowCount;
                }
            }

            internal bool AppendFrom(in Cell cell, bool isDate1904)
            {
                return type switch
                {
                    NativeColumnType.String => AppendString(in cell),
                    NativeColumnType.Int64 => Append(_longs, ExcelCellReaders.Parsable(in cell, isDate1904, CultureInfo.InvariantCulture, out long i64), i64),
                    NativeColumnType.Float64 => Append(_doubles, ExcelCellReaders.Parsable(in cell, isDate1904, CultureInfo.InvariantCulture, out double f64), f64),
                    NativeColumnType.Bool => Append(_bools, ExcelCellReaders.Bool(in cell, isDate1904, CultureInfo.InvariantCulture, out bool flag), (byte)(flag ? 1 : 0)),
                    NativeColumnType.Date => AppendDate(in cell, isDate1904),
                    NativeColumnType.Time => AppendTime(in cell, isDate1904),
                    _ => AppendTimestamp(in cell, isDate1904), // NativeColumnType.Timestamp; range already validated
                };
            }

            // Cell.GetString's own stack buffer for the format-a-number branch. Matched here so an
            // unformattable number lands on the same empty result it would have through GetString.
            private const int NumberFormatMaxBytes = 32;

            private bool AppendString(in Cell cell)
            {
                // Reading a string column always succeeds, including for an empty cell (-> ""), so it is
                // never null regardless of `nullable`.
                //
                // Cell.TryFormat emits exactly the bytes GetString would have decoded — the number branch
                // formats, every other cell copies Value verbatim, and for a shared string Value already
                // holds the resolved text rather than its index. Going through it drops a string AND a
                // byte array per row per column: this used to decode UTF-8 into a string and immediately
                // encode it back, which was the single largest allocation in this whole path.
                //
                // It also stops sanitizing: the old round trip replaced malformed UTF-8 with U+FFFD,
                // this copies the file's bytes through unchanged, matching what every other read path
                // in this library (xl_next_row, xl_read_all_blob) already hands back.
                //
                // NativeTypedParseBenchmark, Data/65K_Records_Data.xlsb (5 string columns of 65535 rows,
                // Ryzen 7 5700X, .NET 10.0.10, 15 iterations): 34.89 MB -> 23.61 MB allocated and
                // 62.01 ms -> 48.58 ms. The largest single win on this path, and unlike the marshalling
                // cleanups it moves wall clock too, well outside the +/-0.9 ms error.
                int capacity = Math.Max(cell.Value.Length, NumberFormatMaxBytes);
                if (_scratch.Length < capacity)
                {
                    _scratch = new byte[capacity];
                }
                if (!cell.TryFormat(_scratch, out int written))
                {
                    written = 0;
                }
                _stringData.AddRange(_scratch.AsSpan(0, written));
                _stringOffsets.Add(_stringData.Count);
                RecordValidity(valid: true);
                return true;
            }

            private static readonly int UnixEpochDayNumber = new DateOnly(1970, 1, 1).DayNumber;

            private bool AppendDate(in Cell cell, bool isDate1904)
            {
                bool ok = ExcelCellReaders.DateOnlyAuto(in cell, isDate1904, CultureInfo.InvariantCulture, out DateOnly value);
                // DateOnly.DayNumber is already an int and the ABI's Date column is 4-byte, so this
                // accumulates as int end to end — no widening to long and narrowing back on marshal.
                return Append(_ints, ok, value.DayNumber - UnixEpochDayNumber);
            }

            private bool AppendTime(in Cell cell, bool isDate1904)
            {
                bool ok = ExcelCellReaders.TimeOnlyAuto(in cell, isDate1904, CultureInfo.InvariantCulture, out TimeOnly value);
                return Append(_longs, ok, value.ToTimeSpan().Ticks / 10); // 1 tick = 100ns -> /10 = microseconds
            }

            private bool AppendTimestamp(in Cell cell, bool isDate1904)
            {
                bool ok = ExcelCellReaders.DateTimeAuto(in cell, isDate1904, CultureInfo.InvariantCulture, out DateTime value);
                return Append(_longs, ok, (value - DateTime.UnixEpoch).Ticks / 10);
            }

            // One append for every non-string type: a failed conversion is only tolerable on a nullable
            // column, and the slot it still occupies holds default(T) so every column stays row-aligned.
            // T is always a value type, so the JIT/AOT specializes this per instantiation — no boxing.
            private bool Append<T>(ChunkedBuffer<T> target, bool converted, T value) where T : unmanaged
            {
                if (!converted && !nullable)
                {
                    return false;
                }
                target.Add(converted ? value : default);
                RecordValidity(converted);
                return true;
            }

            private void RecordValidity(bool valid)
            {
                if ((_rowCount & 7) == 0)
                {
                    _validity.Add(0); // every eighth row opens a fresh byte, zeroed so only set bits matter
                }
                if (valid)
                {
                    _validity.Last |= (byte)(1 << (_rowCount & 7));
                }
                else
                {
                    _anyNull = true;
                }
                _rowCount++;
            }

            internal NativeColumn Build()
            {
                IntPtr validity = IntPtr.Zero;
                if (_anyNull)
                {
                    // A NULL validity pointer is the ABI's "no nulls in this column" signal, so the
                    // bitmap only crosses the boundary when at least one row actually needs it.
                    validity = CopyToNativeBlock(_validity);
                }
                return type switch
                {
                    NativeColumnType.String => BuildStringColumn(validity),
                    NativeColumnType.Bool => BuildFixedWidthColumn(_bools, validity),
                    NativeColumnType.Float64 => BuildFixedWidthColumn(_doubles, validity),
                    NativeColumnType.Date => BuildFixedWidthColumn(_ints, validity),
                    _ => BuildFixedWidthColumn(_longs, validity), // Int64, Time, Timestamp — all 8-byte
                };
            }

            // Once the element type is known, every non-string column is the same operation: copy the
            // accumulated values into one native block — a single copy, straight from the chunks the
            // rows were appended into. (The `[.. list]` array each type used to build first was a
            // second, full-size copy of the whole column, and at 65K rows x 8 bytes those landed on
            // the LOH. Measured by NativeTypedParseBenchmark on Data/65K_Records_Data.xlsb, 14
            // columns, Ryzen 7 5700X, .NET 10.0.10: managed allocation 42.18 MB -> 34.89 MB, Gen2
            // collections -33%. A GC-pressure win, not a speed one.)
            private NativeColumn BuildFixedWidthColumn<T>(ChunkedBuffer<T> values, IntPtr validity) where T : unmanaged
            {
                IntPtr block = CopyToNativeBlock(values);
                return new NativeColumn { Type = type, Length = RowCount, Values = block, Validity = validity, Data = IntPtr.Zero, DataLen = 0 };
            }

            // The one type that is not a plain BuildFixedWidthColumn: offsets and data share a single
            // block, with Data an interior pointer just past the offsets (see NativeColumn's doc
            // comment), so the two buffers are copied into one allocation rather than two.
            private NativeColumn BuildStringColumn(IntPtr validity)
            {
                int offsetBytes = _stringOffsets.ByteLength;
                int dataBytes = _stringData.ByteLength;
                int total = checked(offsetBytes + dataBytes);
                IntPtr block = Marshal.AllocHGlobal(total);
                Span<byte> destination = new((void*)block, total);
                _stringOffsets.CopyTo(destination);
                _stringData.CopyTo(destination[offsetBytes..]);
                return new NativeColumn { Type = type, Length = RowCount, Values = block, Validity = validity, Data = IntPtr.Add(block, offsetBytes), DataLen = dataBytes };
            }
        }
    }
}
