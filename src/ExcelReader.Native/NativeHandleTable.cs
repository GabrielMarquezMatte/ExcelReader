using System.Collections.Concurrent;

namespace ExcelReader.Native
{
    /// <summary>
    /// Maps the opaque handle values handed to C callers onto their <see cref="NativeHandle"/>.
    /// </summary>
    /// <remarks>
    /// The value a caller receives is an id from a monotonic counter, never a GCHandle or any other
    /// pointer. That is the whole point: GCHandle table slots and heap addresses are both RECYCLED,
    /// so a stale value can silently start naming a different, live workbook — a double xl_close
    /// would then free somebody else's handle. An id retired by <see cref="TryUnregister"/> is never
    /// handed out again, so a stale handle stays invalid permanently.
    /// </remarks>
    internal static class NativeHandleTable
    {
        private static readonly ConcurrentDictionary<nint, NativeHandle> Live = new();
        private static long _counter;

        internal static nint Register(NativeHandle handle)
        {
            while (true)
            {
                // 0 is the ABI's null handle and must never be issued. The TryAdd guard also makes
                // this correct if the counter ever wrapped (nint is 32-bit on a 32-bit runtime).
                nint id = (nint)Interlocked.Increment(ref _counter);
                if (id != 0 && Live.TryAdd(id, handle))
                {
                    return id;
                }
            }
        }

        internal static NativeHandle? Resolve(nint id)
        {
            return id != 0 && Live.TryGetValue(id, out NativeHandle? handle) ? handle : null;
        }

        internal static bool TryUnregister(nint id, out NativeHandle? handle)
        {
            if (id != 0 && Live.TryRemove(id, out NativeHandle? removed))
            {
                handle = removed;
                return true;
            }

            handle = null;
            return false;
        }
    }
}
