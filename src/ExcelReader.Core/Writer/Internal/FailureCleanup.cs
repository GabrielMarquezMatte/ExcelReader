using System.Diagnostics.CodeAnalysis;

namespace ExcelReader.Core.Writer.Internal
{
    /// <summary>
    /// Releases a resource while an earlier failure is already propagating. That failure is the one the
    /// caller sees; a second error from tearing down the same broken output would only hide it.
    /// </summary>
    internal static class FailureCleanup
    {
        [SuppressMessage("Design", "CA1031:Do not catch general exception types",
            Justification = "Only called from a catch block that rethrows the original failure.")]
        internal static void Dispose(IDisposable? resource)
        {
            try
            {
                resource?.Dispose();
            }
            catch (Exception)
            {
            }
        }

        [SuppressMessage("Design", "CA1031:Do not catch general exception types",
            Justification = "Only called from a catch block that rethrows the original failure.")]
        internal static async ValueTask DisposeAsync(IAsyncDisposable? resource)
        {
            try
            {
                if (resource is not null)
                {
                    await resource.DisposeAsync().ConfigureAwait(false);
                }
            }
            catch (Exception)
            {
            }
        }
    }
}
