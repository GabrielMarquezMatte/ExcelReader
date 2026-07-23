namespace ExcelReader.Core.Reader
{
    // Shared refill plumbing for concrete XLSX/XLSB/CSV enumerators. MoveNext stays concrete in each
    // format; this base only owns the per-buffer operations and pooled row storage.
    /// <summary>Base class holding the pooled buffer and refill plumbing shared by every concrete format's row enumerator.</summary>
    public abstract class PooledStreamRowEnumerator
    {
        private protected Stream? _source;
        private protected readonly CancellationToken _ct;
        private protected readonly BufferedStreamCursor _io;
        private protected readonly CellAccumulator _acc;
        private protected byte[] _buf => _io.Buf;
        private protected int _pos { get => _io.Pos; set => _io.Pos = value; }
        private protected int _len => _io.Len;
        private protected bool _eof => _io.Eof;

        private protected PooledStreamRowEnumerator(Stream source, int maxCellBytes, string limitName, int initialCapacity, CancellationToken ct)
        {
            _source = source;
            _ct = ct;
            _io = new BufferedStreamCursor(maxCellBytes, limitName, initialCapacity);
            _acc = new CellAccumulator(maxCellBytes, limitName);
        }

        private protected void Fill()
        {
            _io.Fill(_source!);
        }

        private protected ValueTask FillAsync()
        {
            return _io.FillAsync(_source!, _ct);
        }

        private protected void Ensure(int count)
        {
            _io.Ensure(_source!, count);
        }

        private protected ValueTask EnsureAsync(int count)
        {
            return _io.EnsureAsync(_source!, count, _ct);
        }

        private protected void ReturnBuffers()
        {
            _io.Return();
            _acc.Return();
        }
    }
}
