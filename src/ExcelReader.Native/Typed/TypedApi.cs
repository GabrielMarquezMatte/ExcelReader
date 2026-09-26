using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Runtime.InteropServices;
using ExcelReader.Core.Parser;
using ExcelReader.Core.Parser.Internal;
using ExcelReader.Core.Reader;
using ExcelReader.Core.Reader.Schema;

namespace ExcelReader.Native.Typed
{
    internal static unsafe partial class TypedApi
    {
        internal static int ParseTyped(NativeHandle? handle, NativeColumnSpec[] specs, int headerRow, out NativeTable table)
        {
            return ParseTypedTable(handle, specs, headerRow, 1, "xl_parse_typed", out table);
        }

        private static int ParseSequential(NativeHandle? handle, NativeColumnSpec[] specs, int headerRow, string cause, out NativeTable table)
        {
            table = default;
            int status = TypedParseSession.OpenTransient(handle, specs, headerRow, cause, out TypedParseSession? session);
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
                NativeApi.SetLastError(exception.Message);
                table = default;
                return NativeStatus.Error;
            }
        }

        internal static NativeTable BuildEmptyTable(NativeColumnSpec[] specs)
        {
            return BuildTable(NewBuilders(specs));
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
                FreeColumn(table.ColumnAt(index));
            }
            FreeBlock(table.Columns);
            table = default;
        }

        private static void FreeColumn(in NativeColumn column)
        {
            FreeBlock(column.Values);
            FreeBlock(column.Validity);
            FreeBlock(column.Data);
        }

        [ThreadStatic]
        internal static Func<int, IntPtr>? AllocOverride;

        [ThreadStatic]
        internal static Action<IntPtr>? FreeOverride;

        private static IntPtr AllocBlock(int byteLength)
        {
            return AllocOverride is { } alloc ? alloc(byteLength) : Marshal.AllocHGlobal(byteLength);
        }

        private static void FreeBlock(IntPtr block)
        {
            if (block == IntPtr.Zero)
            {
                return;
            }
            if (FreeOverride is { } free)
            {
                free(block);
                return;
            }
            Marshal.FreeHGlobal(block);
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
            return BuildTable([builders]);
        }

        private static NativeTable BuildTable(ReadOnlySpan<ColumnBuilder[]> parts)
        {
            int columnCount = parts[0].Length;
            long rowCount = 0;
            if (columnCount > 0)
            {
                foreach (ColumnBuilder[] part in parts)
                {
                    rowCount += part[0].RowCount;
                }
            }
            IntPtr columnsBlock = AllocBlock(checked(columnCount * sizeof(NativeColumn)));
            NativeColumn* columns = (NativeColumn*)columnsBlock;
            ColumnBuilder[] column = new ColumnBuilder[parts.Length];
            int built = 0;
            try
            {
                for (; built < columnCount; built++)
                {
                    for (int p = 0; p < parts.Length; p++)
                    {
                        column[p] = parts[p][built];
                    }
                    columns[built] = ColumnBuilder.Build(column);
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

        internal static IntPtr PackBitsLsbFirst(ReadOnlySpan<byte> flags)
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

            private int Type
            {
                get
                {
                    return type;
                }
            }

            internal bool AppendFrom(in Cell cell, bool isDate1904)
            {
                return type switch
                {
                    NativeColumnType.String => AppendString(in cell),
                    NativeColumnType.Int64 => Append(_longs, ExcelCellReaders.Parsable(in cell, isDate1904, CultureInfo.InvariantCulture, out long i64), i64),
                    NativeColumnType.Float64 => Append(_doubles, ExcelCellReaders.Parsable(in cell, isDate1904, CultureInfo.InvariantCulture, out double f64), f64),
                    NativeColumnType.Bool => Append(_bools, ColumnParserFactory.ReadBool(in cell, isDate1904, CultureInfo.InvariantCulture, out bool flag), (byte)(flag ? 1 : 0)),
                    NativeColumnType.Date => AppendDate(in cell, isDate1904),
                    NativeColumnType.Time => AppendTime(in cell, isDate1904),
                    _ => AppendTimestamp(in cell, isDate1904),
                };
            }

            private const int NumberFormatMaxBytes = 32;

            private bool AppendString(in Cell cell)
            {
                if (!cell.Value.IsEmpty)
                {
                    _stringData.AddRange(cell.Value);
                    _stringOffsets.Add(_stringData.Count);
                    RecordValidity(valid: true);
                    return true;
                }
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
                bool ok = ColumnParserFactory.ReadDateOnly(in cell, isDate1904, CultureInfo.InvariantCulture, out DateOnly value)
                    || ColumnParserFactory.ReadTextDateOnly(in cell, isDate1904, CultureInfo.InvariantCulture, out value);
                return Append(_ints, ok, value.DayNumber - UnixEpochDayNumber);
            }

            private bool AppendTime(in Cell cell, bool isDate1904)
            {
                bool ok = ColumnParserFactory.ReadTimeOnly(in cell, isDate1904, CultureInfo.InvariantCulture, out TimeOnly value)
                    || ColumnParserFactory.ReadTextTimeOnly(in cell, isDate1904, CultureInfo.InvariantCulture, out value);
                return Append(_longs, ok, value.ToTimeSpan().Ticks / 10);
            }

            private bool AppendTimestamp(in Cell cell, bool isDate1904)
            {
                bool ok = ColumnParserFactory.ReadDateTime(in cell, isDate1904, CultureInfo.InvariantCulture, out DateTime value)
                    || ColumnParserFactory.ReadTextDateTime(in cell, isDate1904, CultureInfo.InvariantCulture, out value);
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

            /// <summary>
            /// Tracks validity without touching the bitmap until the first null: a column with no nulls
            /// ships no bitmap, so writing one bit per cell up to then is wasted work. The first null
            /// backfills every earlier row as valid.
            /// </summary>
            private void RecordValidity(bool valid)
            {
                if (!_anyNull)
                {
                    if (valid)
                    {
                        _rowCount++;
                        return;
                    }
                    _anyNull = true;
                    for (int i = 0; i < _rowCount >> 3; i++)
                    {
                        _validity.Add(0xFF);
                    }
                    if ((_rowCount & 7) != 0)
                    {
                        _validity.Add((byte)((1 << (_rowCount & 7)) - 1));
                    }
                }
                if ((_rowCount & 7) == 0)
                {
                    _validity.Add(0);
                }
                if (valid)
                {
                    _validity.Last |= (byte)(1 << (_rowCount & 7));
                }
                _rowCount++;
            }

            internal static NativeColumn Build(ReadOnlySpan<ColumnBuilder> parts)
            {
                int type = parts[0].Type;
                long rows = 0;
                bool anyNull = false;
                foreach (ColumnBuilder part in parts)
                {
                    rows += part._rowCount;
                    anyNull |= part._anyNull;
                }
                NativeColumn column = new() { Type = type, Length = rows };
                try
                {
                    if (anyNull)
                    {
                        column.Validity = ConcatValidity(parts, rows);
                    }
                    switch (type)
                    {
                        case NativeColumnType.String:
                            column.Values = ConcatStringOffsets(parts);
                            column.Data = Concat(parts, static part => part._stringData, out int dataLength);
                            column.DataLen = dataLength;
                            break;
                        case NativeColumnType.Bool:
                            column.Values = Concat(parts, static part => part._bools, out _);
                            break;
                        case NativeColumnType.Float64:
                            column.Values = Concat(parts, static part => part._doubles, out _);
                            break;
                        case NativeColumnType.Date:
                            column.Values = Concat(parts, static part => part._ints, out _);
                            break;
                        default:
                            column.Values = Concat(parts, static part => part._longs, out _);
                            break;
                    }
                }
                catch
                {
                    FreeColumn(in column);
                    throw;
                }
                return column;
            }

            private static IntPtr Concat<T>(ReadOnlySpan<ColumnBuilder> parts, Func<ColumnBuilder, ChunkedBuffer<T>> select, out int byteLength)
                where T : unmanaged
            {
                byteLength = 0;
                foreach (ColumnBuilder part in parts)
                {
                    byteLength = checked(byteLength + select(part).ByteLength);
                }
                IntPtr block = AllocBlock(Math.Max(byteLength, 1));
                int offset = 0;
                foreach (ColumnBuilder part in parts)
                {
                    ChunkedBuffer<T> buffer = select(part);
                    buffer.CopyTo(new Span<byte>((byte*)block + offset, buffer.ByteLength));
                    offset += buffer.ByteLength;
                }
                return block;
            }

            private static IntPtr ConcatStringOffsets(ReadOnlySpan<ColumnBuilder> parts)
            {
                int count = 1;
                foreach (ColumnBuilder part in parts)
                {
                    count = checked(count + part._stringOffsets.Count - 1);
                }
                IntPtr block = AllocBlock(checked(count * sizeof(int)));
                Span<int> offsets = new((void*)block, count);
                offsets[0] = 0;
                int at = 0;
                foreach (ColumnBuilder part in parts)
                {
                    int baseOffset = offsets[at];
                    Span<int> slot = offsets.Slice(at, part._stringOffsets.Count);
                    part._stringOffsets.CopyTo(MemoryMarshal.AsBytes(slot));
                    if (baseOffset != 0)
                    {
                        foreach (ref int offset in slot)
                        {
                            offset += baseOffset;
                        }
                    }
                    at += slot.Length - 1;
                }
                return block;
            }

            private static IntPtr ConcatValidity(ReadOnlySpan<ColumnBuilder> parts, long rows)
            {
                int byteLength = checked((int)((rows + 7) / 8));
                IntPtr block = AllocBlock(Math.Max(byteLength, 1));
                Span<byte> bits = new((void*)block, byteLength);
                bits.Clear();
                long at = 0;
                foreach (ColumnBuilder part in parts)
                {
                    if (part._anyNull)
                    {
                        OrBits(bits, at, part._validity);
                    }
                    else
                    {
                        SetBits(bits, at, part._rowCount);
                    }
                    at += part._rowCount;
                }
                return block;
            }

            private static void OrBits(Span<byte> bits, long at, ChunkedBuffer<byte> validity)
            {
                byte[] source = new byte[validity.Count];
                validity.CopyTo(source);
                int first = (int)(at >> 3);
                int shift = (int)(at & 7);
                if (shift == 0)
                {
                    source.CopyTo(bits[first..]);
                    return;
                }
                for (int i = 0; i < source.Length; i++)
                {
                    bits[first + i] |= (byte)(source[i] << shift);
                    if (first + i + 1 < bits.Length)
                    {
                        bits[first + i + 1] |= (byte)(source[i] >> (8 - shift));
                    }
                }
            }

            private static void SetBits(Span<byte> bits, long start, int count)
            {
                long end = start + count;
                for (; start < end && (start & 7) != 0; start++)
                {
                    bits[(int)(start >> 3)] |= (byte)(1 << (int)(start & 7));
                }
                int fullBytes = (int)((end - start) >> 3);
                bits.Slice((int)(start >> 3), fullBytes).Fill(0xFF);
                for (start += (long)fullBytes << 3; start < end; start++)
                {
                    bits[(int)(start >> 3)] |= (byte)(1 << (int)(start & 7));
                }
            }
        }
    }
}
