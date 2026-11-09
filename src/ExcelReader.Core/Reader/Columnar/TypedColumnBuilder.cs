using System.Runtime.CompilerServices;
using ExcelReader.Core.Parser;
using ExcelReader.Core.Parser.Internal;
using ExcelReader.Core.Reader.Internal;
using ExcelReader.Core.Reader.Schema;

namespace ExcelReader.Core.Reader.Columnar
{
    internal sealed partial class TypedColumnBuilder(TypedColumnSpec spec) : IDisposable
    {
        private const int NumberFormatMaxBytes = 32;
        private static readonly int UnixEpochDayNumber = new DateOnly(1970, 1, 1).DayNumber;

        private readonly ExcelColumnType _type = spec.Type;
        private readonly bool _nullable = spec.Nullable;
        private readonly bool _isText = spec.Type is ExcelColumnType.StringColumn or ExcelColumnType.StringDictionaryColumn;

        private readonly GrowableNativeBuffer<byte> _validity = new();
        private int _rowCount;
        private bool _anyNull;

        private readonly GrowableNativeBuffer<long> _longs = new();
        private readonly GrowableNativeBuffer<int> _ints = new();
        private readonly GrowableNativeBuffer<double> _doubles = new();
        private readonly GrowableNativeBuffer<byte> _bools = new();
        private readonly GrowableNativeBuffer<int> _stringOffsets = NewStringOffsets(spec.Type);
        private readonly GrowableNativeBuffer<byte> _stringData = new();
        private readonly StringDictionary? _dictionary = spec.Type == ExcelColumnType.StringDictionaryColumn ? new StringDictionary() : null;
        private int[] _sharedEntries = [];

        private byte[] _scratch = [];

        internal int RowCount => _rowCount;

        internal ExcelColumnType Type => _type;

        public void Dispose()
        {
            _validity.Dispose();
            _longs.Dispose();
            _ints.Dispose();
            _doubles.Dispose();
            _bools.Dispose();
            _stringOffsets.Dispose();
            _stringData.Dispose();
            _dictionary?.Dispose();
        }

        private static GrowableNativeBuffer<int> NewStringOffsets(ExcelColumnType type)
        {
            GrowableNativeBuffer<int> offsets = new();
            if (type == ExcelColumnType.StringColumn)
            {
                offsets.Add(0);
            }
            return offsets;
        }

        internal bool AppendFrom(in Cell cell, bool isDate1904)
        {
            return _type switch
            {
                ExcelColumnType.StringColumn => AppendString(in cell),
                ExcelColumnType.StringDictionaryColumn => AppendDictionaryString(in cell),
                ExcelColumnType.Int64Column => AppendInt64(in cell, isDate1904),
                ExcelColumnType.Float64Column => AppendFloat64(in cell, isDate1904),
                ExcelColumnType.BoolColumn => AppendBool(in cell, isDate1904),
                ExcelColumnType.DateColumn => AppendDate(in cell, isDate1904),
                ExcelColumnType.TimeColumn => AppendTime(in cell, isDate1904),
                _ => AppendTimestamp(in cell, isDate1904),
            };
        }

        // The common cases read straight from the cell's descriptor, with the same result AppendFrom
        // gives them: text bytes as they are, a binary number as Cell.TryParse<double/long> returns it.
        // Everything else returns false and goes through AppendFrom.
        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        internal bool CanAppendDirect(in CellDesc cell)
        {
            return cell.HasNumber || _isText;
        }

        internal bool TryAppendDirect(in CellDesc cell, in Row row)
        {
            if (_type == ExcelColumnType.StringColumn && cell.Length > 0)
            {
                _stringData.AddRange(row.ValueOf(in cell));
                _stringOffsets.Add(_stringData.Count);
                RecordValidity(valid: true);
                return true;
            }
            if (!cell.HasNumber)
            {
                return TryAppendDictionaryEntry(in cell, in row);
            }
            double number = cell.Number;
            if (_type == ExcelColumnType.Float64Column)
            {
                _doubles.Add(number);
                RecordValidity(valid: true);
                return true;
            }
            if (_type == ExcelColumnType.Int64Column && number == Math.Truncate(number) && number is >= -9223372036854775808.0 and < 9223372036854775808.0)
            {
                _longs.Add((long)number);
                RecordValidity(valid: true);
                return true;
            }
            return false;
        }

        private bool TryAppendDictionaryEntry(in CellDesc cell, in Row row)
        {
            if (_type != ExcelColumnType.StringDictionaryColumn || cell.Length == 0)
            {
                return false;
            }
            _ints.Add(DictionaryEntry(in cell, in row));
            RecordValidity(valid: true);
            return true;
        }

        // A shared string's index finds its entry without hashing the bytes again.
        private int DictionaryEntry(in CellDesc cell, in Row row)
        {
            if (cell.Source != CellValueSource.Shared || cell.SharedIndex < 0)
            {
                return _dictionary!.GetOrAdd(row.ValueOf(in cell));
            }
            int shared = cell.SharedIndex;
            if (shared < _sharedEntries.Length && _sharedEntries[shared] != 0)
            {
                return _sharedEntries[shared] - 1;
            }
            int entry = _dictionary!.GetOrAdd(row.ValueOf(in cell));
            if (shared >= _sharedEntries.Length)
            {
                Array.Resize(ref _sharedEntries, Math.Max(shared + 1, _sharedEntries.Length * 2));
            }
            _sharedEntries[shared] = entry + 1;
            return entry;
        }

        private bool AppendString(in Cell cell)
        {
            if (IsBlankNull(in cell))
            {
                return AppendNull(_stringOffsets, _stringData.Count);
            }
            _stringData.AddRange(TextOf(in cell));
            _stringOffsets.Add(_stringData.Count);
            RecordValidity(valid: true);
            return true;
        }

        private bool AppendDictionaryString(in Cell cell)
        {
            if (IsBlankNull(in cell))
            {
                return AppendNull(_ints, 0);
            }
            _ints.Add(_dictionary!.GetOrAdd(TextOf(in cell)));
            RecordValidity(valid: true);
            return true;
        }

        private bool IsBlankNull(in Cell cell)
        {
            return _nullable && cell.Type == CellType.Empty;
        }

        private ReadOnlySpan<byte> TextOf(in Cell cell)
        {
            if (!cell.Value.IsEmpty)
            {
                return cell.Value;
            }
            if (_scratch.Length < NumberFormatMaxBytes)
            {
                _scratch = new byte[NumberFormatMaxBytes];
            }
            if (!cell.TryFormat(_scratch, out int written))
            {
                return [];
            }
            return _scratch.AsSpan(0, written);
        }

        private bool AppendNull<T>(GrowableNativeBuffer<T> target, T filler) where T : unmanaged
        {
            target.Add(filler);
            RecordValidity(valid: false);
            return true;
        }

        [MethodImpl(MethodImplOptions.NoInlining)]
        private bool AppendInt64(in Cell cell, bool isDate1904)
        {
            return Append(_longs, ExcelCellReaders.Parsable(in cell, new ExcelRowContext { IsDate1904 = isDate1904 }, out long value), value);
        }

        [MethodImpl(MethodImplOptions.NoInlining)]
        private bool AppendFloat64(in Cell cell, bool isDate1904)
        {
            return Append(_doubles, ExcelCellReaders.Parsable(in cell, new ExcelRowContext { IsDate1904 = isDate1904 }, out double value), value);
        }

        [MethodImpl(MethodImplOptions.NoInlining)]
        private bool AppendBool(in Cell cell, bool isDate1904)
        {
            return Append(_bools, ColumnParserFactory.ReadBool(in cell, new ExcelRowContext { IsDate1904 = isDate1904 }, out bool flag), (byte)(flag ? 1 : 0));
        }

        private bool AppendDate(in Cell cell, bool isDate1904)
        {
            bool ok = ColumnParserFactory.ReadDateOnly(in cell, new ExcelRowContext { IsDate1904 = isDate1904 }, out DateOnly value)
                || ColumnParserFactory.ReadTextDateOnly(in cell, new ExcelRowContext { IsDate1904 = isDate1904 }, out value);
            return Append(_ints, ok, value.DayNumber - UnixEpochDayNumber);
        }

        private bool AppendTime(in Cell cell, bool isDate1904)
        {
            bool ok = ColumnParserFactory.ReadTimeOnly(in cell, new ExcelRowContext { IsDate1904 = isDate1904 }, out TimeOnly value)
                || ColumnParserFactory.ReadTextTimeOnly(in cell, new ExcelRowContext { IsDate1904 = isDate1904 }, out value);
            return Append(_longs, ok, value.ToTimeSpan().Ticks / 10);
        }

        private bool AppendTimestamp(in Cell cell, bool isDate1904)
        {
            bool ok = ColumnParserFactory.ReadDateTime(in cell, new ExcelRowContext { IsDate1904 = isDate1904 }, out DateTime value)
                || ColumnParserFactory.ReadTextDateTime(in cell, new ExcelRowContext { IsDate1904 = isDate1904 }, out value);
            // Text with a "Z" or an offset parses to the machine's local time; the wire value is UTC.
            if (value.Kind == DateTimeKind.Local)
            {
                value = value.ToUniversalTime();
            }
            return Append(_longs, ok, (value - DateTime.UnixEpoch).Ticks / 10);
        }

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        private bool Append<T>(GrowableNativeBuffer<T> target, bool converted, T value) where T : unmanaged
        {
            if (!converted && !_nullable)
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
        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        private void RecordValidity(bool valid)
        {
            if (!_anyNull && valid)
            {
                _rowCount++;
                return;
            }
            RecordValiditySlow(valid);
        }

        [MethodImpl(MethodImplOptions.NoInlining)]
        private void RecordValiditySlow(bool valid)
        {
            if (!_anyNull)
            {
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
    }
}
