using System.Diagnostics.CodeAnalysis;
using ExcelReader.Core.Writer.Internal;

namespace ExcelReader.Core.Writer
{
    public sealed class RowWriter : IAsyncDisposable
    {
        [SuppressMessage("SharpSource", "SS066:DisposableFieldIsNotDisposed",
            Justification = "SheetWriter is borrowed; its lifetime is managed by the caller.")]
        private readonly SheetWriter _owner;
        [SuppressMessage("SharpSource", "SS066:DisposableFieldIsNotDisposed",
            Justification = "StreamWriter is owned by SheetWriter; RowWriter borrows it.")]
        private readonly StreamWriter _xml;
        private readonly int _rowNumber;
        private int _columnIndex;
        private bool _disposed;

        internal RowWriter(SheetWriter owner, StreamWriter xml, int rowNumber)
        {
            _owner = owner;
            _xml = xml;
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
                CellFormatter.WriteEmpty(_xml, _columnIndex, _rowNumber);
                _columnIndex++;
                return;
            }
            CellFormatter.WriteString(_xml, value, _columnIndex, _rowNumber);
            _columnIndex++;
        }

        public void Write(bool value)
        {
            ThrowIfDisposed();
            CellFormatter.WriteBool(_xml, value, _columnIndex, _rowNumber);
            _columnIndex++;
        }

        public void Write(bool? value)
        {
            ThrowIfDisposed();
            if (value is null)
            {
                CellFormatter.WriteEmpty(_xml, _columnIndex, _rowNumber);
                _columnIndex++;
                return;
            }
            CellFormatter.WriteBool(_xml, value.Value, _columnIndex, _rowNumber);
            _columnIndex++;
        }

        public void Write(DateTime value)
        {
            ThrowIfDisposed();
            CellFormatter.WriteDateTime(_xml, value, _columnIndex, _rowNumber);
            _columnIndex++;
        }

        public void Write(DateTime? value)
        {
            ThrowIfDisposed();
            if (value is null)
            {
                CellFormatter.WriteEmpty(_xml, _columnIndex, _rowNumber);
                _columnIndex++;
                return;
            }
            CellFormatter.WriteDateTime(_xml, value.Value, _columnIndex, _rowNumber);
            _columnIndex++;
        }

        public void Write<T>(T value)
            where T : ISpanFormattable
        {
            ThrowIfDisposed();
            CellFormatter.WriteNumber(_xml, value, _columnIndex, _rowNumber);
            _columnIndex++;
        }

        public void Write<T>(T? value)
            where T : struct, ISpanFormattable
        {
            ThrowIfDisposed();
            if (value is null)
            {
                CellFormatter.WriteEmpty(_xml, _columnIndex, _rowNumber);
                _columnIndex++;
                return;
            }
            CellFormatter.WriteNumber(_xml, value.Value, _columnIndex, _rowNumber);
            _columnIndex++;
        }

        public void Skip(int count = 1)
        {
            ThrowIfDisposed();
            _columnIndex += count;
        }

        public async ValueTask DisposeAsync()
        {
            if (_disposed)
            {
                return;
            }
            _disposed = true;
            await _xml.WriteAsync("</row>").ConfigureAwait(false);
            _owner.NotifyRowEnded();
        }
    }
}
