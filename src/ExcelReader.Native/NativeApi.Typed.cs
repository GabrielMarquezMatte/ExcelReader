using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Runtime.InteropServices;
using ExcelReader.Core.Parser;
using ExcelReader.Core.Reader;
using ExcelReader.Core.ValueObjects;

namespace ExcelReader.Native
{
    internal static unsafe partial class NativeApi
    {
        internal static int ParseTyped(NativeHandle? handle, NativeColumnSpec[] specs, int headerRow, out NativeTable table)
        {
            table = default;
            int status = TypedParseSession.OpenTransient(handle, specs, headerRow, "xl_parse_typed",
                out TypedParseSession? session);
            if (status != NativeStatus.Ok)
            {
                return status;
            }

            using TypedParseSession open = session!;
            status = open.NextBatch(out table);
            if (status != NativeStatus.Eof)
            {
                return status;
            }

            try
            {
                table = BuildEmptyTable(specs);
                return NativeStatus.Ok;
            }
            catch (Exception exception)
            {
                SetLastError(exception.Message);
                table = default;
                return NativeStatus.Error;
            }
        }

        private static NativeTable BuildEmptyTable(NativeColumnSpec[] specs)
        {
            ColumnBuilder[] builders = new ColumnBuilder[specs.Length];
            for (int i = 0; i < specs.Length; i++)
            {
                builders[i] = new ColumnBuilder(specs[i].Type, specs[i].Nullable);
            }
            return BuildTable(builders);
        }

        private static string DescribeFailedColumn(NativeColumnSpec[] specs, int failedColumn)
        {
            NativeColumnSpec spec = specs[failedColumn];
            string columnLabel = spec.Names.Length > 0
                ? string.Join(" / ", spec.Names)
                : spec.Index.ToString(CultureInfo.InvariantCulture);
            return $"column {failedColumn} (\"{columnLabel}\") has a value that failed to convert and is not nullable.";
        }

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

        internal static bool IsValidSpecCount(int specCount)
        {
            return specCount is > 0 and <= NativeLimits.MaxColumnSpecs;
        }

        internal static bool IsValidNameLength(int nameLength)
        {
            return nameLength is >= 0 and <= NativeLimits.MaxColumnNameBytes;
        }

        internal static bool IsValidNameCount(int nameCount)
        {
            return nameCount is >= 0 and <= NativeLimits.MaxNamesPerSpec;
        }

        private static bool TryValidateArguments(NativeColumnSpec[] specs, int headerRow, [NotNullWhen(false)] out string? error)
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

        private static bool TryResolveColumns(IExcelRowEnumerator rows, NativeColumnSpec[] specs, int headerRow, int[] columnIndices, [NotNullWhen(false)] out string? error)
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
            int built = 0;
            try
            {
                for (; built < columnCount; built++)
                {
                    columns[built] = builders[built].Build();
                }
            }
            catch
            {
                NativeTable partial = new() { ColumnCount = built, RowCount = rowCount, Columns = columnsBlock };
                FreeTable(ref partial);
                throw;
            }
            return new NativeTable { ColumnCount = columnCount, RowCount = rowCount, Columns = columnsBlock };
        }

        private static IntPtr CopyToNativeBlock<T>(ChunkedBuffer<T> source) where T : unmanaged
        {
            int byteLength = source.ByteLength;
            IntPtr block = Marshal.AllocHGlobal(Math.Max(byteLength, 1));
            source.CopyTo(new Span<byte>((void*)block, byteLength));
            return block;
        }

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

        private sealed class ColumnBuilder(int type, bool nullable)
        {
            private readonly ChunkedBuffer<byte> _validity = new();
            private int _rowCount;
            private bool _anyNull;

            private readonly ChunkedBuffer<long> _longs = new();
            private readonly ChunkedBuffer<int> _ints = new();
            private readonly ChunkedBuffer<double> _doubles = new();
            private readonly ChunkedBuffer<byte> _bools = new();
            private readonly ChunkedBuffer<int> _stringOffsets = NewStringOffsets(type);
            private readonly ChunkedBuffer<byte> _stringData = new();

            private static ChunkedBuffer<int> NewStringOffsets(int type)
            {
                ChunkedBuffer<int> offsets = new();
                if (type == NativeColumnType.String)
                {
                    offsets.Add(0);
                }
                return offsets;
            }

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
                    _ => AppendTimestamp(in cell, isDate1904),
                };
            }

            private const int NumberFormatMaxBytes = 32;

            private bool AppendString(in Cell cell)
            {
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
                return Append(_ints, ok, value.DayNumber - UnixEpochDayNumber);
            }

            private bool AppendTime(in Cell cell, bool isDate1904)
            {
                bool ok = ExcelCellReaders.TimeOnlyAuto(in cell, isDate1904, CultureInfo.InvariantCulture, out TimeOnly value);
                return Append(_longs, ok, value.ToTimeSpan().Ticks / 10);
            }

            private bool AppendTimestamp(in Cell cell, bool isDate1904)
            {
                bool ok = ExcelCellReaders.DateTimeAuto(in cell, isDate1904, CultureInfo.InvariantCulture, out DateTime value);
                return Append(_longs, ok, (value - DateTime.UnixEpoch).Ticks / 10);
            }

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
                    _validity.Add(0);
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
                    validity = CopyToNativeBlock(_validity);
                }
                return type switch
                {
                    NativeColumnType.String => BuildStringColumn(validity),
                    NativeColumnType.Bool => BuildFixedWidthColumn(_bools, validity),
                    NativeColumnType.Float64 => BuildFixedWidthColumn(_doubles, validity),
                    NativeColumnType.Date => BuildFixedWidthColumn(_ints, validity),
                    _ => BuildFixedWidthColumn(_longs, validity),
                };
            }

            private NativeColumn BuildFixedWidthColumn<T>(ChunkedBuffer<T> values, IntPtr validity) where T : unmanaged
            {
                IntPtr block = CopyToNativeBlock(values);
                return new NativeColumn { Type = type, Length = RowCount, Values = block, Validity = validity, Data = IntPtr.Zero, DataLen = 0 };
            }

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
