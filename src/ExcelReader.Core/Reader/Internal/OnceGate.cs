using System.Runtime.ExceptionServices;

namespace ExcelReader.Core.Reader.Internal
{
    /// <summary>
    /// Runs a load once for however many callers ask. A failure is kept and rethrown; a cancelled
    /// load is undone so the next caller starts clean.
    /// </summary>
    internal sealed class OnceGate : IDisposable
    {
        private readonly SemaphoreSlim _gate = new(1, 1);
        private volatile bool _done;
        private ExceptionDispatchInfo? _failure;

        internal bool IsDone => _done;

        internal void Run<TState>(TState state, Action<TState> load, Action<TState> reset)
        {
            if (_done)
            {
                return;
            }
            _gate.Wait();
            try
            {
                if (_done)
                {
                    return;
                }
                _failure?.Throw();
                try
                {
                    load(state);
                }
                catch (OperationCanceledException)
                {
                    reset(state);
                    throw;
                }
                catch (Exception ex)
                {
                    _failure = ExceptionDispatchInfo.Capture(ex);
                    throw;
                }
                _done = true;
            }
            finally
            {
                _gate.Release();
            }
        }

        internal async ValueTask RunAsync<TState>(
            TState state, Func<TState, CancellationToken, ValueTask> load, Action<TState> reset, CancellationToken ct)
        {
            if (_done)
            {
                return;
            }
            await _gate.WaitAsync(ct).ConfigureAwait(false);
            try
            {
                if (_done)
                {
                    return;
                }
                _failure?.Throw();
                try
                {
                    await load(state, ct).ConfigureAwait(false);
                }
                catch (OperationCanceledException)
                {
                    reset(state);
                    throw;
                }
                catch (Exception ex)
                {
                    _failure = ExceptionDispatchInfo.Capture(ex);
                    throw;
                }
                _done = true;
            }
            finally
            {
                _gate.Release();
            }
        }

        public void Dispose()
        {
            _gate.Dispose();
        }
    }
}
