namespace ExcelReader.Core.Parser.Internal
{
    // A fixed set of recycled chunk model lists, shared by the workers and the merge of one
    // enumeration.
    //
    // Not ArrayPool and not ConcurrentStack: the population is tiny and known up front (the
    // semaphore already bounds how many lists can exist at once), so a flat array of slots swapped
    // with Interlocked beats both — no node allocation per return, and no cross-enumeration pool
    // holding parsed rows alive after the enumeration that produced them has finished.
    //
    // A miss is not an error: Rent allocates, Return drops. That keeps the merge and the workers off
    // any blocking path, at the cost of falling back to exactly the behaviour this type replaces.
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

        // The caller clears the list before handing it back, so a slot never holds parsed rows.
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
