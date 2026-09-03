using System.Diagnostics.CodeAnalysis;

namespace ExcelReader.Core.Reader
{
    /// <summary>Base class holding the pooled buffer and refill plumbing shared by every concrete format's row enumerator.</summary>
    /// <remarks><c>MoveNext</c>/<c>MoveNextAsync</c> stay concrete in each derived format; this base only owns the per-buffer operations and pooled row storage.</remarks>
    public abstract class PooledStreamRowEnumerator
    {
        private protected Stream? _source;
        private protected readonly CancellationToken _ct;
        private protected readonly BufferedStreamCursor _io;
        private protected readonly CellAccumulator _acc;
        private readonly bool _ownsSource;
        private protected byte[] _buf => _io.Buf;
        private protected int _pos { get => _io.Pos; set => _io.Pos = value; }
        private protected int _len => _io.Len;
        private protected bool _eof => _io.Eof;

        private protected PooledStreamRowEnumerator(
            Stream source, int maxCellBytes, string limitName, int initialCapacity, bool ownsSource, CancellationToken ct)
        {
            _source = source;
            _ct = ct;
            _ownsSource = ownsSource;
            _io = new BufferedStreamCursor(maxCellBytes, limitName, initialCapacity);
            _acc = new CellAccumulator(maxCellBytes, limitName);
        }

        private protected PooledStreamRowEnumerator(ReadOnlyMemory<byte> content, int maxCellBytes, string limitName, CancellationToken ct)
        {
            _source = null;
            _ct = ct;
            _ownsSource = false;
            _io = new BufferedStreamCursor(content, maxCellBytes, limitName);
            _acc = new CellAccumulator(maxCellBytes, limitName);
        }

        /// <summary>Releases the sheet stream (when owned) and returns both pooled buffers to the pool.</summary>
#pragma warning disable S2953 // Methods named "Dispose" should implement "IDisposable.Dispose"
        public void Dispose()
#pragma warning restore S2953 // Methods named "Dispose" should implement "IDisposable.Dispose"
        {
            if (_ownsSource)
            {
                _source?.Dispose();
            }
            _source = null;
            ReturnBuffers();
        }

        /// <summary>Asynchronous counterpart to <see cref="Dispose"/>.</summary>
        public async ValueTask DisposeAsync()
        {
            if (_ownsSource && _source is not null)
            {
                await _source.DisposeAsync().ConfigureAwait(false);
            }
            _source = null;
            ReturnBuffers();
        }

        private protected void Fill()
        {
            _io.Fill(_source);
        }

        private protected ValueTask FillAsync()
        {
            return _io.FillAsync(_source, _ct);
        }

        private protected void Ensure(int count)
        {
            _io.Ensure(_source, count);
        }

        private protected ValueTask EnsureAsync(int count)
        {
            return _io.EnsureAsync(_source, count, _ct);
        }

        private protected void ReturnBuffers()
        {
            _io.Return();
            _acc.Return();
        }
    }
}
