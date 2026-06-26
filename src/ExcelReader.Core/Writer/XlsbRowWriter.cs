using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using ExcelReader.Core.Reader;
using ExcelReader.Core.Writer.Internal;

namespace ExcelReader.Core.Writer
{
    public sealed class XlsbRowWriter : IRowWriter
    {
        [SuppressMessage("SharpSource", "SS066:DisposableFieldIsNotDisposed",
            Justification = "XlsbSheetWriter is borrowed; its lifetime is managed by the caller.")]
        private readonly XlsbSheetWriter _owner;
        [SuppressMessage("SharpSource", "SS066:DisposableFieldIsNotDisposed",
            Justification = "BiffBuffer is owned by XlsbSheetWriter; row writer borrows it.")]
        private readonly BiffBuffer _records;
        private readonly bool _date1904;
        private readonly BiffBuffer _payload = new(256);
        private int _columnIndex;
        private bool _disposed;

        internal XlsbRowWriter(XlsbSheetWriter owner, BiffBuffer records, bool date1904)
        {
            _owner = owner;
            _records = records;
            _date1904 = date1904;
        }

        private void ThrowIfDisposed()
        {
            ObjectDisposedException.ThrowIf(_disposed, this);
        }

        public void Write(string? value)
        {
            ThrowIfDisposed();
            if (value is not null)
            {
                _payload.Reset();
                Biff12RecordWriter.WriteCellHeader(_payload, _columnIndex, 0);
                Biff12RecordWriter.WriteWideString(_payload, value);
                Biff12RecordWriter.WriteRecord(_records, Brt.CellSt, _payload.Span);
            }
            _columnIndex++;
        }

        public void Write(bool value)
        {
            ThrowIfDisposed();
            _payload.Reset();
            Biff12RecordWriter.WriteCellHeader(_payload, _columnIndex, 0);
            _payload.WriteByte(value ? (byte)1 : (byte)0);
            Biff12RecordWriter.WriteRecord(_records, Brt.CellBool, _payload.Span);
            _columnIndex++;
        }

        public void Write(bool? value)
        {
            ThrowIfDisposed();
            if (value is not null)
            {
                Write(value.Value);
                return;
            }
            _columnIndex++;
        }

        public void Write(DateTime value)
        {
            ThrowIfDisposed();
            double serial = value.ToOADate();
            if (_date1904)
            {
                serial -= 1462.0;
            }
            WriteDouble(serial, style: 1);
        }

        public void Write(DateTime? value)
        {
            ThrowIfDisposed();
            if (value is not null)
            {
                Write(value.Value);
                return;
            }
            _columnIndex++;
        }

        public void Write<T>(T value)
            where T : ISpanFormattable
        {
            ThrowIfDisposed();
            WriteDouble(ToDouble(value), style: 0);
        }

        public void Write<T>(T? value)
            where T : struct, ISpanFormattable
        {
            ThrowIfDisposed();
            if (value is not null)
            {
                Write(value.Value);
                return;
            }
            _columnIndex++;
        }

        public void Skip(int count = 1)
        {
            ThrowIfDisposed();
            ArgumentOutOfRangeException.ThrowIfNegative(count);
            _columnIndex += count;
        }

        private void WriteDouble(double value, int style)
        {
            _payload.Reset();
            Biff12RecordWriter.WriteCellHeader(_payload, _columnIndex, style);
            _payload.WriteDouble(value);
            Biff12RecordWriter.WriteRecord(_records, Brt.CellReal, _payload.Span);
            _columnIndex++;
        }

        private static double ToDouble<T>(T value)
            where T : ISpanFormattable
        {
            if (value is IConvertible convertible)
            {
                return convertible.ToDouble(CultureInfo.InvariantCulture);
            }
            Span<char> chars = stackalloc char[64];
            return value.TryFormat(chars, out int written, default, CultureInfo.InvariantCulture) &&
                double.TryParse(chars[..written], CultureInfo.InvariantCulture, out double parsed)
                ? parsed
                : 0.0;
        }

        public ValueTask DisposeAsync()
        {
            if (_disposed)
            {
                return ValueTask.CompletedTask;
            }
            _disposed = true;
            _payload.Dispose();
            _owner.NotifyRowEnded();
            return ValueTask.CompletedTask;
        }
    }
}