using System.Collections.Concurrent;

namespace ExcelReader.Native
{
    internal static class NativeHandleTable
    {
        private static readonly ConcurrentDictionary<nint, object> _live = new();
        private static long _counter;

        internal static nint Register(object handle)
        {
            while (true)
            {
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
