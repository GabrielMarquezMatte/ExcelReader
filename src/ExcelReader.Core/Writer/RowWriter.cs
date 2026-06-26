using System.Diagnostics.CodeAnalysis;
using System.Text;
using ExcelReader.Core.Writer.Internal;

namespace ExcelReader.Core.Writer
{
    public sealed class RowWriter : IRowWriter
    {
        [SuppressMessage("SharpSource", "SS066:DisposableFieldIsNotDisposed",
            Justification = "SheetWriter is borrowed; its lifetime is managed by the caller.")]
        private readonly SheetWriter _owner;
        [SuppressMessage("SharpSource", "SS066:DisposableFieldIsNotDisposed",
            Justification = "StreamWriter is owned by SheetWriter; RowWriter borrows it.")]
        private readonly StreamWriter _xml;
        private readonly StringBuilder _row;
        private readonly int _rowNumber;
        private int _columnIndex;
        private bool _useCellReferences;
        private bool _disposed;

        internal RowWriter(SheetWriter owner, StreamWriter xml, StringBuilder row, int rowNumber)
        {
            _owner = owner;
            _xml = xml;
            _row = row;
            _rowNumber = rowNumber;
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
            _row.Append("</row>");
            _xml.Write(_row);
            _owner.NotifyRowEnded();
            return ValueTask.CompletedTask;
        }
    }
}
