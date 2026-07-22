using System.Diagnostics.CodeAnalysis;
using ExcelReader.Core.Writer.Internal;
using ExcelReader.Core.Reader;

namespace ExcelReader.Core.Writer
{
    public sealed class XlsxRowWriter : IRowWriter, IDisposable
    {
        [SuppressMessage("SharpSource", "SS066:DisposableFieldIsNotDisposed",
            Justification = "XlsxSheetWriter is borrowed; its lifetime is managed by the caller.")]
        private readonly XlsxSheetWriter _owner;
        [SuppressMessage("SharpSource", "SS066:DisposableFieldIsNotDisposed",
            Justification = "BiffBuffer is owned by XlsxSheetWriter; XlsxRowWriter borrows it.")]
        private readonly BiffBuffer _row;
        private int _rowNumber;
        private int _columnIndex;
        private bool _useCellReferences;
        private bool _disposed;

        internal XlsxRowWriter(XlsxSheetWriter owner, BiffBuffer row)
        {
            _owner = owner;
            _row = row;
        }

        // Reused across rows by XlsxSheetWriter: rents one instance per sheet instead of one per row.
        internal void Reset(int rowNumber)
        {
            _rowNumber = rowNumber;
            _columnIndex = 0;
            _useCellReferences = false;
            _disposed = false;
        }

        private void ThrowIfDisposed()
        {
            ObjectDisposedException.ThrowIf(_disposed, this);
        }

        public void Write(string? value)
        {
            ThrowIfDisposed();
            if (value is null)
            {
                WriteEmptyCell();
                return;
            }
            // Excel's universal per-cell text limit, enforced here so every writer rejects the same way
            // instead of each format silently truncating or corrupting its own record encoding.
            const int maxCellTextLength = 32_767;
            if (value.Length > maxCellTextLength)
            {
                throw new ArgumentException(
                    $"Cell text exceeds Excel's {maxCellTextLength}-character limit ({value.Length} chars).", nameof(value));
            }
            if (_owner.UseSharedStrings)
            {
                int index = _owner.GetSharedStringIndex(value);
                CellFormatter.WriteSharedString(_row, index, _columnIndex, _rowNumber, ConsumeCellReference());
                _columnIndex++;
                return;
            }
            CellFormatter.WriteString(_row, value, _columnIndex, _rowNumber, ConsumeCellReference());
            _columnIndex++;
        }

        public void Write(bool value)
        {
            ThrowIfDisposed();
            CellFormatter.WriteBool(_row, value, _columnIndex, _rowNumber, ConsumeCellReference());
            _columnIndex++;
        }

        public void Write(bool? value)
        {
            ThrowIfDisposed();
            if (value is null)
            {
                WriteEmptyCell();
                return;
            }
            Write(value.Value);
        }

        public void Write(DateTime value)
        {
            ThrowIfDisposed();
            CellFormatter.WriteDateTime(_row, value, _columnIndex, _rowNumber, ConsumeCellReference());
            _columnIndex++;
        }

        public void Write(DateTime? value)
        {
            ThrowIfDisposed();
            if (value is null)
            {
                WriteEmptyCell();
                return;
            }
            Write(value.Value);
        }

        // DateOnly shares the DateTime date-serial cell format (midnight), so it round-trips as a date.
        public void Write(DateOnly value)
        {
            ThrowIfDisposed();
            CellFormatter.WriteDateTime(_row, value.ToDateTime(TimeOnly.MinValue), _columnIndex, _rowNumber, ConsumeCellReference());
            _columnIndex++;
        }

        public void Write(DateOnly? value)
        {
            ThrowIfDisposed();
            if (value is null)
            {
                WriteEmptyCell();
                return;
            }
            Write(value.Value);
        }

        // TimeOnly is written as an Excel time serial: the fraction of a 24h day, in [0,1). This is a
        // plain number cell (no date style); ExcelParser<T> reconstructs the TimeOnly from the fraction.
        public void Write(TimeOnly value)
        {
            ThrowIfDisposed();
            CellFormatter.WriteNumber(_row, value.Ticks / (double)TimeSpan.TicksPerDay, _columnIndex, _rowNumber, ConsumeCellReference());
            _columnIndex++;
        }

        public void Write(TimeOnly? value)
        {
            ThrowIfDisposed();
            if (value is null)
            {
                WriteEmptyCell();
                return;
            }
            Write(value.Value);
        }

        public void Write(int value)
        {
            ThrowIfDisposed();
            CellFormatter.WriteNumber(_row, value, _columnIndex, _rowNumber, ConsumeCellReference());
            _columnIndex++;
        }

        public void Write(int? value)
        {
            ThrowIfDisposed();
            if (value is null)
            {
                WriteEmptyCell();
                return;
            }
            Write(value.Value);
        }

        public void Write(long value)
        {
            ThrowIfDisposed();
            CellFormatter.WriteNumber(_row, value, _columnIndex, _rowNumber, ConsumeCellReference());
            _columnIndex++;
        }

        public void Write(long? value)
        {
            ThrowIfDisposed();
            if (value is null)
            {
                WriteEmptyCell();
                return;
            }
            Write(value.Value);
        }

        public void Write(double value)
        {
            ThrowIfDisposed();
            CellFormatter.WriteNumber(_row, value, _columnIndex, _rowNumber, ConsumeCellReference());
            _columnIndex++;
        }

        public void Write(double? value)
        {
            ThrowIfDisposed();
            if (value is null)
            {
                WriteEmptyCell();
                return;
            }
            Write(value.Value);
        }

        public void Write(decimal value)
        {
            ThrowIfDisposed();
            CellFormatter.WriteNumber(_row, value, _columnIndex, _rowNumber, ConsumeCellReference());
            _columnIndex++;
        }

        public void Write(decimal? value)
        {
            ThrowIfDisposed();
            if (value is null)
            {
                WriteEmptyCell();
                return;
            }
            Write(value.Value);
        }

        public void Write<T>(T value)
            where T : IUtf8SpanFormattable
        {
            ThrowIfDisposed();
            CellFormatter.WriteNumber(_row, value, _columnIndex, _rowNumber, ConsumeCellReference());
            _columnIndex++;
        }

        public void Write<T>(T? value)
            where T : struct, IUtf8SpanFormattable
        {
            ThrowIfDisposed();
            if (value is null)
            {
                WriteEmptyCell();
                return;
            }
            Write(value.Value);
        }

        public void Skip(int count = 1)
        {
            ThrowIfDisposed();
            ArgumentOutOfRangeException.ThrowIfNegative(count);
            if (count > 0)
            {
                if (_columnIndex > 16_384 - count)
                {
                    throw new ExcelLimitExceededException("Columns", 16_384, (long)_columnIndex + count);
                }
                _useCellReferences = true;
            }
            _columnIndex += count;
        }

        // Only the first cell written after a gap (a real Skip(), which leaves the corresponding
        // <c> elements out entirely) needs an explicit r="..." reference so the reader can tell how
        // many columns were skipped; every cell after that is contiguous again. A null-valued cell
        // still emits a real (self-closing) <c/> placeholder, so on its own it never needs one.
        private bool ConsumeCellReference()
        {
            if ((uint)_columnIndex >= 16_384)
            {
                throw new ExcelLimitExceededException("Columns", 16_384, _columnIndex + 1L);
            }
            bool result = _useCellReferences;
            _useCellReferences = false;
            return result;
        }

        private void WriteEmptyCell()
        {
            CellFormatter.WriteEmpty(_row, _columnIndex, _rowNumber, includeReference: ConsumeCellReference());
            _columnIndex++;
        }

        public ValueTask DisposeAsync()
        {
            if (_disposed)
            {
                return ValueTask.CompletedTask;
            }
            _disposed = true;
            return _owner.EndBufferedRowAsync();
        }

        // Sync counterpart for callers on XlsxSheetWriter.StartRow's synchronous fast path (see
        // SheetWriterExtensions.WriteRecordsAsync's XlsxSheetWriter-specific overload): avoids the
        // per-row await entirely when the destination stream write only happens on the rare
        // buffer-threshold flush, which EndBufferedRow already does synchronously.
        public void Dispose()
        {
            if (_disposed)
            {
                return;
            }
            _disposed = true;
            _owner.EndBufferedRow();
        }
    }
}
