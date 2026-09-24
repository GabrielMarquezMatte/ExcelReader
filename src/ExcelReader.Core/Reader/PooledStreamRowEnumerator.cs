using System.ComponentModel;
using System.Diagnostics;
using ExcelReader.Core.Reader.Internal;

namespace ExcelReader.Core.Reader
{
    /// <summary>Base class holding the pooled buffer and refill plumbing shared by every concrete format's row enumerator.</summary>
    /// <remarks><c>MoveNext</c>/<c>MoveNextAsync</c> stay concrete in each derived format; this base only owns the per-buffer operations and pooled row storage.</remarks>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public abstract class PooledStreamRowEnumerator
    {
        private protected Stream? _source;
        private protected readonly CancellationToken _ct;
        private protected readonly BufferedStreamCursor _io;
        private protected readonly CellAccumulator _acc;
        private readonly bool _ownsSource;
        private bool _deferred;
        private protected byte[] _buf => _io.Buf;
        private protected int _pos { get => _io.Pos; set => _io.Pos = value; }
        private protected int _len => _io.Len;
        private protected bool _eof => _io.Eof;

        private protected PooledStreamRowEnumerator(
            Stream? source, int maxCellBytes, string limitName, int initialCapacity, bool ownsSource, CancellationToken ct)
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

        /// <summary>Deferred form: the source is opened by <see cref="OpenSource"/>/<see cref="OpenSourceAsync"/> on the first refill.</summary>
        private protected PooledStreamRowEnumerator(int maxCellBytes, string limitName, int initialCapacity, CancellationToken ct)
            : this(source: null, maxCellBytes, limitName, initialCapacity, ownsSource: true, ct)
        {
            _deferred = true;
        }

        private protected virtual Stream OpenSource()
        {
            throw new UnreachableException();
        }

        private protected virtual ValueTask<Stream> OpenSourceAsync()
        {
            throw new UnreachableException();
        }

        private void OpenDeferred()
        {
            _source = OpenSource();
            _deferred = false;
        }

        private async ValueTask OpenDeferredAsync()
        {
            _source = await OpenSourceAsync().ConfigureAwait(false);
            _deferred = false;
        }

        /// <summary>Releases the sheet stream (when owned) and returns both pooled buffers to the pool.</summary>
#pragma warning disable S2953
        public void Dispose()
#pragma warning restore S2953 
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
            if (_deferred)
            {
                OpenDeferred();
            }
            _io.Fill(_source);
        }

        private protected ValueTask FillAsync()
        {
            return _deferred ? OpenThenFillAsync() : _io.FillAsync(_source, _ct);
        }

        private async ValueTask OpenThenFillAsync()
        {
            await OpenDeferredAsync().ConfigureAwait(false);
            await _io.FillAsync(_source, _ct).ConfigureAwait(false);
        }

        private protected void Ensure(int count)
        {
            if (_io.Len - _io.Pos < count && !_io.Eof)
            {
                EnsureSlow(count);
            }
        }

        private void EnsureSlow(int count)
        {
            if (_deferred)
            {
                OpenDeferred();
            }
            _io.Ensure(_source, count);
        }

        private protected ValueTask EnsureAsync(int count)
        {
            if (_io.Len - _io.Pos >= count || _io.Eof)
            {
                return ValueTask.CompletedTask;
            }
            return EnsureSlowAsync(count);
        }

        private ValueTask EnsureSlowAsync(int count)
        {
            return _deferred ? OpenThenEnsureAsync(count) : _io.EnsureAsync(_source, count, _ct);
        }

        private async ValueTask OpenThenEnsureAsync(int count)
        {
            await OpenDeferredAsync().ConfigureAwait(false);
            await _io.EnsureAsync(_source, count, _ct).ConfigureAwait(false);
        }

        private protected virtual void ReturnBuffers()
        {
            _io.Return();
            _acc.Return();
        }
    }
}
