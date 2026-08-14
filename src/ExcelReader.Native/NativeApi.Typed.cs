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
                        SetLastError($"column {failedColumn} (\"{spec.Name ?? spec.Index.ToString(CultureInfo.InvariantCulture)}\") has a value that failed to convert and is not nullable.");
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
                if (spec.Name is null && spec.Index < 0)
                {
                    error = "a column spec with no name must have a non-negative index.";
                    return false;
                }
                if (spec.Name is not null && headerRow == 0)
                {
                    error = $"column \"{spec.Name}\" is name-based, but header_row is 0 (no header row to match it against).";
                    return false;
                }
                if (spec.Type is < NativeColumnType.String or > NativeColumnType.Timestamp)
                {
                    error = $"column spec has unknown type {spec.Type}.";
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

            for (int rowNumber = 1; rowNumber <= headerRow; rowNumber++)
            {
                if (!rows.MoveNext())
                {
                    error = $"sheet has fewer than {headerRow} row(s); cannot resolve header_row.";
                    return false;
                }
            }

            Row header = rows.Current;
            for (int i = 0; i < specs.Length; i++)
            {
                if (specs[i].Name is not string name)
                {
                    columnIndices[i] = specs[i].Index;
                    continue;
                }
                int found = FindHeaderColumn(header, name);
                if (found < 0)
                {
                    error = $"no column header matches \"{name}\".";
                    return false;
                }
                columnIndices[i] = found;
            }
            return true;
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

        // A validity bitmap and Arrow's canonical boolean layout are the same thing — one LSB-first bit
        // per row in a native block sized (n + 7) / 8 — so both go through here; the callers differ only
        // in where their one-byte-per-row source lives.
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
            private readonly List<byte> _validity = []; // 1 = valid, 0 = null; one entry per row
            private bool _anyNull;

            // Only one of these is populated, chosen by `type` — see AppendFrom.
            private readonly List<long> _longs = []; // Int64, Date (truncated to int32 on marshal), Time, Timestamp
            private readonly List<double> _doubles = []; // Float64
            private readonly List<byte> _bools = []; // Bool, one byte (0/1) per row
            private readonly List<int> _stringOffsets = [0]; // String
            private readonly List<byte> _stringData = []; // String

            internal int RowCount
            {
                get
                {
                    return _validity.Count;
                }
            }

            internal bool AppendFrom(in Cell cell, bool isDate1904)
            {
                return type switch
                {
                    NativeColumnType.String => AppendString(in cell),
                    NativeColumnType.Int64 => Append(_longs, ExcelCellReaders.Parsable<long>(in cell, isDate1904, CultureInfo.InvariantCulture, out long i64), i64),
                    NativeColumnType.Float64 => Append(_doubles, ExcelCellReaders.Parsable<double>(in cell, isDate1904, CultureInfo.InvariantCulture, out double f64), f64),
                    NativeColumnType.Bool => Append(_bools, ExcelCellReaders.Bool(in cell, isDate1904, CultureInfo.InvariantCulture, out bool flag), (byte)(flag ? 1 : 0)),
                    NativeColumnType.Date => AppendDate(in cell, isDate1904),
                    NativeColumnType.Time => AppendTime(in cell, isDate1904),
                    _ => AppendTimestamp(in cell, isDate1904), // NativeColumnType.Timestamp; range already validated
                };
            }

            private bool AppendString(in Cell cell)
            {
                // ExcelCellReaders.String (cell.GetString()) always succeeds, including for an empty
                // cell (-> ""), so a string column is never null regardless of `nullable`.
                byte[] utf8 = Encoding.UTF8.GetBytes(cell.GetString());
                _stringData.AddRange(utf8);
                _stringOffsets.Add(_stringData.Count);
                _validity.Add(1);
                return true;
            }

            private static readonly int UnixEpochDayNumber = new DateOnly(1970, 1, 1).DayNumber;

            private bool AppendDate(in Cell cell, bool isDate1904)
            {
                bool ok = ExcelCellReaders.DateOnlyAuto(in cell, isDate1904, CultureInfo.InvariantCulture, out DateOnly value);
                // Excel's own date range (1900-9999) is nowhere near int32's day-count range, so this
                // narrowing is always exact for real data - no `checked` needed here.
                return Append(_longs, ok, value.DayNumber - UnixEpochDayNumber);
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
            private bool Append<T>(List<T> target, bool converted, T value) where T : struct
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
                _validity.Add(valid ? (byte)1 : (byte)0);
                _anyNull |= !valid;
            }

            internal NativeColumn Build()
            {
                IntPtr validity = IntPtr.Zero;
                if (_anyNull)
                {
                    validity = PackBitsLsbFirst(CollectionsMarshal.AsSpan(_validity));
                }
                return type switch
                {
                    NativeColumnType.String => BuildStringColumn(validity),
                    NativeColumnType.Bool => BuildFixedWidthColumn(CollectionsMarshal.AsSpan(_bools), validity),
                    NativeColumnType.Float64 => BuildFixedWidthColumn(CollectionsMarshal.AsSpan(_doubles), validity),
                    NativeColumnType.Date => BuildDateColumn(validity),
                    _ => BuildFixedWidthColumn(CollectionsMarshal.AsSpan(_longs), validity), // Int64, Time, Timestamp — all 8-byte, unlike Date
                };
            }

            // Once the element type is known, every non-string column is the same operation: copy the
            // accumulated values into one native block. Reading the backing List<T> as a span keeps that
            // at a single copy — the `[.. list]` array each type used to build first was a second,
            // full-size copy of the whole column, and at 65K rows x 8 bytes those landed on the LOH.
            // Measured by NativeTypedParseBenchmark on Data/65K_Records_Data.xlsb (14 columns, Ryzen 7
            // 5700X, .NET 10.0.10): managed allocation 42.18 MB -> 34.89 MB, Gen2 collections -33%.
            // Wall clock did not move outside the noise, so this is a GC-pressure win, not a speed one.
            private NativeColumn BuildFixedWidthColumn<T>(ReadOnlySpan<T> values, IntPtr validity) where T : unmanaged
            {
                ReadOnlySpan<byte> source = MemoryMarshal.AsBytes(values);
                IntPtr block = Marshal.AllocHGlobal(Math.Max(source.Length, 1));
                source.CopyTo(new Span<byte>((void*)block, source.Length));
                return new NativeColumn { Type = type, Length = RowCount, Values = block, Validity = validity, Data = IntPtr.Zero, DataLen = 0 };
            }

            // Excel's own date range (1900-9999) is nowhere near int32's day-count range, so the
            // narrowing from _longs (accumulated as long for every non-Float64/Bool/String type) is
            // always exact for real data - no `checked` needed here.
            private NativeColumn BuildDateColumn(IntPtr validity)
            {
                int[] values = new int[_longs.Count];
                for (int i = 0; i < values.Length; i++)
                {
                    values[i] = unchecked((int)_longs[i]);
                }
                return BuildFixedWidthColumn(values, validity);
            }

            // The one type that is not a plain BuildFixedWidthColumn: offsets and data share a single
            // block, with Data an interior pointer just past the offsets (see NativeColumn's doc
            // comment), so the two spans are copied into one allocation rather than two.
            private NativeColumn BuildStringColumn(IntPtr validity)
            {
                ReadOnlySpan<byte> offsets = MemoryMarshal.AsBytes(CollectionsMarshal.AsSpan(_stringOffsets));
                ReadOnlySpan<byte> data = CollectionsMarshal.AsSpan(_stringData);
                int total = checked(offsets.Length + data.Length);
                IntPtr block = Marshal.AllocHGlobal(total);
                Span<byte> destination = new((void*)block, total);
                offsets.CopyTo(destination);
                data.CopyTo(destination[offsets.Length..]);
                return new NativeColumn { Type = type, Length = RowCount, Values = block, Validity = validity, Data = IntPtr.Add(block, offsets.Length), DataLen = data.Length };
            }
        }
    }
}
