namespace ExcelReader.Core.Reader.Internal
{
    /// <summary>
    /// Keeps a reader's resources alive while any enumerator still uses them: the reader holds one
    /// reference, each enumerator another, and the release runs when the last one goes.
    /// </summary>
    internal sealed class ReaderLifetime(Action release)
    {
        private int _count = 1;
        private int _closed;

        internal void Acquire(object owner)
        {
            while (true)
            {
                int current = Volatile.Read(ref _count);
                ObjectDisposedException.ThrowIf(current == 0 || Volatile.Read(ref _closed) != 0, owner);
                if (Interlocked.CompareExchange(ref _count, current + 1, current) == current)
                {
                    return;
                }
            }
        }

        internal void ThrowIfClosed(object owner)
        {
            ObjectDisposedException.ThrowIf(Volatile.Read(ref _closed) != 0, owner);
        }

        internal void Release()
        {
            if (Interlocked.Decrement(ref _count) == 0)
            {
                release();
            }
        }

        internal void Close()
        {
            if (Interlocked.Exchange(ref _closed, 1) == 0)
            {
                Release();
            }
        }
    }
}
