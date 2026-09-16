namespace ExcelReader.Core.Parser.Internal
{
    internal sealed class ListPool<T>
    {
        private readonly List<T>?[] _slots;

        internal ListPool(int capacity)
        {
            _slots = new List<T>?[capacity];
        }

        internal List<T> Rent()
        {
            foreach (ref List<T>? slot in _slots.AsSpan())
            {
                if (slot is null)
                {
                    continue;
                }
                List<T>? taken = Interlocked.Exchange(ref slot, null);
                if (taken is not null)
                {
                    return taken;
                }
            }
            return [];
        }

        internal void Return(List<T> list)
        {
            foreach (ref List<T>? slot in _slots.AsSpan())
            {
                if (slot is null && Interlocked.CompareExchange(ref slot, list, null) is null)
                {
                    return;
                }
            }
        }
    }
}
