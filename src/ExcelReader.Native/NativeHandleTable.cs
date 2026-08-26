using System.Collections.Concurrent;

namespace ExcelReader.Native
{
    /// <summary>
    /// Maps the opaque handle values handed to C callers onto their live managed object — a
    /// <see cref="NativeHandle"/> (reader) or a <see cref="Writer.NativeWriterHandle"/> (writer).
    /// </summary>
    /// <remarks>
    /// The value a caller receives is an id from a single monotonic counter shared by every handle
    /// kind, never a GCHandle or any other pointer. That is the whole point: GCHandle table slots and
    /// heap addresses are both RECYCLED, so a stale value can silently start naming a different, live
    /// object — a double xl_close would then free somebody else's handle. An id retired by
    /// <see cref="TryUnregister{T}"/> is never handed out again, so a stale handle stays invalid
    /// permanently.
    /// </remarks>
    /// <remarks>
    /// The counter and table are shared across every handle kind rather than one per kind: reader ids
    /// and writer ids must never collide, since the ABI hands out a bare <c>nint</c> with no type tag.
    /// <see cref="Resolve{T}"/>/<see cref="TryUnregister{T}"/> additionally check the stored object's
    /// runtime type, so passing a live writer id to a reader entry point (or vice versa) resolves to
    /// nothing rather than to the wrong kind of handle.
    /// </remarks>
    internal static class NativeHandleTable
    {
        private static readonly ConcurrentDictionary<nint, object> _live = new();
        private static long _counter;

        internal static nint Register(object handle)
        {
            while (true)
            {
                // 0 is the ABI's null handle and must never be issued. The TryAdd guard also makes
                // this correct if the counter ever wrapped (nint is 32-bit on a 32-bit runtime).
                nint id = (nint)Interlocked.Increment(ref _counter);
                if (id != 0 && _live.TryAdd(id, handle))
                {
                    return id;
                }
            }
        }

        internal static T? Resolve<T>(nint id) where T : class
        {
            return id != 0 && _live.TryGetValue(id, out object? handle) ? handle as T : null;
        }

        internal static bool TryUnregister<T>(nint id, out T? handle) where T : class
        {
            // Removed via the exact (id, value) pair, not TryRemove(id, out _): a plain id-only
            // removal would also match a live object of the WRONG kind (e.g. a writer id passed to
            // the reader's xl_close), retiring it under the wrong caller's control. Checking `is T`
            // first (in the id's own read) and then removing that exact pair keeps a cross-kind call
            // a clean no-op instead of stealing another kind's handle out from under it.
            if (id != 0 && _live.TryGetValue(id, out object? existing) && existing is T typed
                && _live.TryRemove(new KeyValuePair<nint, object>(id, existing)))
            {
                handle = typed;
                return true;
            }

            handle = null;
            return false;
        }
    }
}
