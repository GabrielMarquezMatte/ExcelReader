using System.Globalization;
using System.Runtime.InteropServices;
using System.Text;
using ExcelReader.Core.Enums;
using ExcelReader.Core.Parser;
using ExcelReader.Core.Reader;
using ExcelReader.Core.ValueObjects;

namespace ExcelReader.Native
{
    internal static partial class NativeApi
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

                List<ColumnBuilder> builders = [.. specs.Select(spec => new ColumnBuilder(spec.Type, spec.Nullable))];
                bool isDate1904 = handle.Reader.IsDate1904;
                while (rows.MoveNext())
                {
                    Row row = rows.Current;
                    for (int i = 0; i < builders.Count; i++)
                    {
                        if (!builders[i].AppendFrom(row[columnIndices[i]], isDate1904))
                        {
                            SetLastError($"column {i} (\"{specs[i].Name ?? specs[i].Index.ToString(CultureInfo.InvariantCulture)}\") has a value that failed to convert and is not nullable.");
                            return NativeStatus.Error;
                        }
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

        /// <summary>Releases a result returned by <see cref="ParseTyped"/> and resets it to zero. Safe on a zeroed value.</summary>
        internal static void FreeTable(ref NativeTable table)
        {
            if (table.Columns == IntPtr.Zero)
            {
                table = default;
                return;
            }

            int columnSize = Marshal.SizeOf<NativeColumn>();
            for (int index = 0; index < table.ColumnCount; index++)
            {
                NativeColumn column = Marshal.PtrToStructure<NativeColumn>(IntPtr.Add(table.Columns, index * columnSize));
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

        // Mirrors ExcelParserConfig.Default's HeaderNormalization.Trim + OrdinalIgnoreCase comparer,
        // reimplemented with public APIs only: HeaderNormalizationExtensions.Apply is internal to
        // ExcelReader.Core, and this project has no InternalsVisibleTo access to it.
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

        private static NativeTable BuildTable(List<ColumnBuilder> builders)
        {
            int columnCount = builders.Count;
            long rowCount = columnCount > 0 ? builders[0].RowCount : 0;
            int columnSize = Marshal.SizeOf<NativeColumn>();
            IntPtr columnsBlock = Marshal.AllocHGlobal(checked(columnCount * columnSize));
            for (int i = 0; i < columnCount; i++)
            {
                Marshal.StructureToPtr(builders[i].Build(), IntPtr.Add(columnsBlock, i * columnSize), false);
            }
            return new NativeTable { ColumnCount = columnCount, RowCount = rowCount, Columns = columnsBlock };
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

            internal int RowCount => _validity.Count;

            internal bool AppendFrom(in Cell cell, bool isDate1904)
            {
                return type switch
                {
                    NativeColumnType.String => AppendString(in cell),
                    NativeColumnType.Int64 => AppendLong(ExcelCellReaders.Parsable<long>(in cell, isDate1904, CultureInfo.InvariantCulture, out long i64), i64),
                    NativeColumnType.Float64 => AppendDouble(ExcelCellReaders.Parsable<double>(in cell, isDate1904, CultureInfo.InvariantCulture, out double f64), f64),
                    NativeColumnType.Bool => AppendBool(ExcelCellReaders.Bool(in cell, isDate1904, CultureInfo.InvariantCulture, out bool flag), flag),
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
                return AppendLong(ok, value.DayNumber - UnixEpochDayNumber);
            }

            private bool AppendTime(in Cell cell, bool isDate1904)
            {
                bool ok = ExcelCellReaders.TimeOnlyAuto(in cell, isDate1904, CultureInfo.InvariantCulture, out TimeOnly value);
                return AppendLong(ok, value.ToTimeSpan().Ticks / 10); // 1 tick = 100ns -> /10 = microseconds
            }

            private bool AppendTimestamp(in Cell cell, bool isDate1904)
            {
                bool ok = ExcelCellReaders.DateTimeAuto(in cell, isDate1904, CultureInfo.InvariantCulture, out DateTime value);
                return AppendLong(ok, (value - DateTime.UnixEpoch).Ticks / 10);
            }

            private bool AppendLong(bool converted, long value)
            {
                if (!converted && !nullable)
                {
                    return false;
                }
                _longs.Add(converted ? value : 0);
                RecordValidity(converted);
                return true;
            }

            private bool AppendDouble(bool converted, double value)
            {
                if (!converted && !nullable)
                {
                    return false;
                }
                _doubles.Add(converted ? value : 0);
                RecordValidity(converted);
                return true;
            }

            private bool AppendBool(bool converted, bool value)
            {
                if (!converted && !nullable)
                {
                    return false;
                }
                _bools.Add((byte)(converted && value ? 1 : 0));
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
                IntPtr validity = _anyNull ? BuildValidityBitmap() : IntPtr.Zero;
                return type switch
                {
                    NativeColumnType.String => BuildStringColumn(validity),
                    NativeColumnType.Bool => BuildBoolColumn(validity),
                    NativeColumnType.Float64 => BuildFloat64Column(validity),
                    NativeColumnType.Date => BuildDateColumn(validity),
                    _ => BuildLongColumn(validity), // Int64, Time, Timestamp — all 8-byte, unlike Date
                };
            }

            private NativeColumn BuildBoolColumn(IntPtr validity)
            {
                byte[] values = [.. _bools];
                IntPtr block = Marshal.AllocHGlobal(Math.Max(values.Length, 1));
                if (values.Length > 0)
                {
                    Marshal.Copy(values, 0, block, values.Length);
                }
                return new NativeColumn { Type = type, Length = RowCount, Values = block, Validity = validity, Data = IntPtr.Zero, DataLen = 0 };
            }

            private NativeColumn BuildFloat64Column(IntPtr validity)
            {
                double[] values = [.. _doubles];
                IntPtr block = Marshal.AllocHGlobal(Math.Max(values.Length * sizeof(double), 1));
                if (values.Length > 0)
                {
                    Marshal.Copy(values, 0, block, values.Length);
                }
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
                IntPtr block = Marshal.AllocHGlobal(Math.Max(values.Length * sizeof(int), 1));
                if (values.Length > 0)
                {
                    Marshal.Copy(values, 0, block, values.Length);
                }
                return new NativeColumn { Type = type, Length = RowCount, Values = block, Validity = validity, Data = IntPtr.Zero, DataLen = 0 };
            }

            private NativeColumn BuildLongColumn(IntPtr validity)
            {
                long[] values = [.. _longs];
                IntPtr block = Marshal.AllocHGlobal(Math.Max(values.Length * sizeof(long), 1));
                if (values.Length > 0)
                {
                    Marshal.Copy(values, 0, block, values.Length);
                }
                return new NativeColumn { Type = type, Length = RowCount, Values = block, Validity = validity, Data = IntPtr.Zero, DataLen = 0 };
            }

            private NativeColumn BuildStringColumn(IntPtr validity)
            {
                int[] offsets = [.. _stringOffsets];
                byte[] data = [.. _stringData];
                int offsetsBytes = offsets.Length * sizeof(int);
                IntPtr block = Marshal.AllocHGlobal(checked(offsetsBytes + data.Length));
                Marshal.Copy(offsets, 0, block, offsets.Length);
                IntPtr dataPtr = IntPtr.Add(block, offsetsBytes);
                if (data.Length > 0)
                {
                    Marshal.Copy(data, 0, dataPtr, data.Length);
                }
                return new NativeColumn { Type = type, Length = RowCount, Values = block, Validity = validity, Data = dataPtr, DataLen = data.Length };
            }

            private IntPtr BuildValidityBitmap()
            {
                int rowCount = _validity.Count;
                byte[] bitmap = new byte[Math.Max((rowCount + 7) / 8, 1)];
                for (int i = 0; i < rowCount; i++)
                {
                    if (_validity[i] != 0)
                    {
                        bitmap[i >> 3] |= (byte)(1 << (i & 7));
                    }
                }
                IntPtr block = Marshal.AllocHGlobal(bitmap.Length);
                Marshal.Copy(bitmap, 0, block, bitmap.Length);
                return block;
            }
        }
    }
}
