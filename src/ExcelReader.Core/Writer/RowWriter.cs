using System.Diagnostics.CodeAnalysis;
using ExcelReader.Core.Writer.Internal;

namespace ExcelReader.Core.Writer
{
    public sealed class RowWriter : IRowWriter
    {
        [SuppressMessage("SharpSource", "SS066:DisposableFieldIsNotDisposed",
            Justification = "SheetWriter is borrowed; its lifetime is managed by the caller.")]
        private readonly SheetWriter _owner;
        [SuppressMessage("SharpSource", "SS066:DisposableFieldIsNotDisposed",
            Justification = "BiffBuffer is owned by SheetWriter; RowWriter borrows it.")]
        private readonly BiffBuffer _row;
        private int _rowNumber;
        private int _columnIndex;
        private bool _useCellReferences;
        private bool _disposed;

        internal RowWriter(SheetWriter owner, BiffBuffer row)
        {
            _owner = owner;
            _row = row;
        }

        // Reused across rows by SheetWriter: rents one instance per sheet instead of one per row.
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
            if (_owner.UseSharedStrings)
            {
                int index = _owner.GetSharedStringIndex(value);
                CellFormatter.WriteSharedString(_row, index, _columnIndex, _rowNumber, _useCellReferences);
                _columnIndex++;
                return;
            }
            CellFormatter.WriteString(_row, value, _columnIndex, _rowNumber, _useCellReferences);
            _columnIndex++;
        }

        public void Write(bool value)
        {
            ThrowIfDisposed();
            CellFormatter.WriteBool(_row, value, _columnIndex, _rowNumber, _useCellReferences);
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
            CellFormatter.WriteBool(_row, value.Value, _columnIndex, _rowNumber, _useCellReferences);
            _columnIndex++;
        }

        public void Write(DateTime value)
        {
            ThrowIfDisposed();
            CellFormatter.WriteDateTime(_row, value, _columnIndex, _rowNumber, _useCellReferences);
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
            CellFormatter.WriteDateTime(_row, value.Value, _columnIndex, _rowNumber, _useCellReferences);
            _columnIndex++;
        }

        // DateOnly shares the DateTime date-serial cell format (midnight), so it round-trips as a date.
        public void Write(DateOnly value)
        {
            ThrowIfDisposed();
            CellFormatter.WriteDateTime(_row, value.ToDateTime(TimeOnly.MinValue), _columnIndex, _rowNumber, _useCellReferences);
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
            CellFormatter.WriteDateTime(_row, value.Value.ToDateTime(TimeOnly.MinValue), _columnIndex, _rowNumber, _useCellReferences);
            _columnIndex++;
        }

        // TimeOnly is written as an Excel time serial: the fraction of a 24h day, in [0,1). This is a
        // plain number cell (no date style); ExcelParser<T> reconstructs the TimeOnly from the fraction.
        public void Write(TimeOnly value)
        {
            ThrowIfDisposed();
            CellFormatter.WriteNumber(_row, value.Ticks / (double)TimeSpan.TicksPerDay, _columnIndex, _rowNumber, _useCellReferences);
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
            CellFormatter.WriteNumber(_row, value.Value.Ticks / (double)TimeSpan.TicksPerDay, _columnIndex, _rowNumber, _useCellReferences);
            _columnIndex++;
        }

        public void Write(int value)
        {
            ThrowIfDisposed();
            CellFormatter.WriteNumber(_row, value, _columnIndex, _rowNumber, _useCellReferences);
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
            CellFormatter.WriteNumber(_row, value.Value, _columnIndex, _rowNumber, _useCellReferences);
            _columnIndex++;
        }

        public void Write(long value)
        {
            ThrowIfDisposed();
            CellFormatter.WriteNumber(_row, value, _columnIndex, _rowNumber, _useCellReferences);
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
            CellFormatter.WriteNumber(_row, value.Value, _columnIndex, _rowNumber, _useCellReferences);
            _columnIndex++;
        }

        public void Write(double value)
        {
            ThrowIfDisposed();
            CellFormatter.WriteNumber(_row, value, _columnIndex, _rowNumber, _useCellReferences);
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
            CellFormatter.WriteNumber(_row, value.Value, _columnIndex, _rowNumber, _useCellReferences);
            _columnIndex++;
        }

        public void Write(decimal value)
        {
            ThrowIfDisposed();
            CellFormatter.WriteNumber(_row, value, _columnIndex, _rowNumber, _useCellReferences);
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
            CellFormatter.WriteNumber(_row, value.Value, _columnIndex, _rowNumber, _useCellReferences);
            _columnIndex++;
        }

        public void Write<T>(T value)
            where T : ISpanFormattable
        {
            ThrowIfDisposed();
            CellFormatter.WriteNumber(_row, value, _columnIndex, _rowNumber, _useCellReferences);
            _columnIndex++;
        }

        public void Write<T>(T? value)
            where T : struct, ISpanFormattable
        {
            ThrowIfDisposed();
            if (value is null)
            {
                WriteEmptyCell();
                return;
            }
            CellFormatter.WriteNumber(_row, value.Value, _columnIndex, _rowNumber, _useCellReferences);
            _columnIndex++;
        }

        public void Skip(int count = 1)
        {
            ThrowIfDisposed();
            if (count > 0)
            {
                _useCellReferences = true;
            }
            _columnIndex += count;
        }

        private void WriteEmptyCell()
        {
            _useCellReferences = true;
            CellFormatter.WriteEmpty(_row, _columnIndex, _rowNumber, includeReference: true);
            _columnIndex++;
        }

        [SuppressMessage("Reliability", "CA1849:Call async methods when in an async method",
            Justification = "Rows are buffered in memory and flushed synchronously to avoid per-row async state machines.")]
        [SuppressMessage("VisualStudio.Threading", "VSTHRD103:Dispose synchronously blocks",
            Justification = "See CA1849 justification above.")]
        [SuppressMessage("SharpSource", "SS033:Async overload available",
            Justification = "See CA1849 justification above.")]
        public ValueTask DisposeAsync()
        {
            if (_disposed)
            {
                return ValueTask.CompletedTask;
            }
            _disposed = true;
            _owner.EndBufferedRow();
            return ValueTask.CompletedTask;
        }
    }
}
