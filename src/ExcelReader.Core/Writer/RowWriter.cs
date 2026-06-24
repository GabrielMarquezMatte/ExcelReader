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

        public ValueTask WriteAsync(string? value, CancellationToken ct = default)
        {
            ThrowIfDisposed();
            ct.ThrowIfCancellationRequested();
            if (value is null)
            {
                CellFormatter.WriteEmpty(_xml, _columnIndex, _rowNumber);
            }
            else
            {
                CellFormatter.WriteString(_xml, value, _columnIndex, _rowNumber);
            }
            _columnIndex++;
            return ValueTask.CompletedTask;
        }

        public ValueTask WriteAsync(bool value, CancellationToken ct = default)
        {
            ThrowIfDisposed();
            ct.ThrowIfCancellationRequested();
            CellFormatter.WriteBool(_xml, value, _columnIndex, _rowNumber);
            _columnIndex++;
            return ValueTask.CompletedTask;
        }

        public ValueTask WriteAsync(bool? value, CancellationToken ct = default)
        {
            ThrowIfDisposed();
            ct.ThrowIfCancellationRequested();
            if (value is null)
            {
                CellFormatter.WriteEmpty(_xml, _columnIndex, _rowNumber);
            }
            else
            {
                CellFormatter.WriteBool(_xml, value.Value, _columnIndex, _rowNumber);
            }
            _columnIndex++;
            return ValueTask.CompletedTask;
        }

        public ValueTask WriteAsync(DateTime value, CancellationToken ct = default)
        {
            ThrowIfDisposed();
            ct.ThrowIfCancellationRequested();
            CellFormatter.WriteDateTime(_xml, value, _columnIndex, _rowNumber);
            _columnIndex++;
            return ValueTask.CompletedTask;
        }

        public ValueTask WriteAsync(DateTime? value, CancellationToken ct = default)
        {
            ThrowIfDisposed();
            ct.ThrowIfCancellationRequested();
            if (value is null)
            {
                CellFormatter.WriteEmpty(_xml, _columnIndex, _rowNumber);
            }
            else
            {
                CellFormatter.WriteDateTime(_xml, value.Value, _columnIndex, _rowNumber);
            }
            _columnIndex++;
            return ValueTask.CompletedTask;
        }

        public ValueTask WriteAsync<T>(T value, CancellationToken ct = default)
            where T : ISpanFormattable
        {
            ThrowIfDisposed();
            ct.ThrowIfCancellationRequested();
            CellFormatter.WriteNumber(_xml, value, _columnIndex, _rowNumber);
            _columnIndex++;
            return ValueTask.CompletedTask;
        }

        public ValueTask WriteAsync<T>(T? value, CancellationToken ct = default)
            where T : struct, ISpanFormattable
        {
            ThrowIfDisposed();
            ct.ThrowIfCancellationRequested();
            if (value is null)
            {
                CellFormatter.WriteEmpty(_xml, _columnIndex, _rowNumber);
            }
            else
            {
                CellFormatter.WriteNumber(_xml, value.Value, _columnIndex, _rowNumber);
            }
            _columnIndex++;
            return ValueTask.CompletedTask;
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
