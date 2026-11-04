namespace ExcelReader.NetFramework.Interop
{
    internal sealed class LoadGuard
    {
        private readonly Lazy<bool> _loaded;

        internal LoadGuard(Action load)
        {
            _loaded = new Lazy<bool>(
                () =>
                {
                    load();
                    return true;
                },
                LazyThreadSafetyMode.ExecutionAndPublication);
        }

        // Lazy caches a failure and rethrows the original exception on every later call.
        internal void Ensure()
        {
            _ = _loaded.Value;
        }
    }
}
