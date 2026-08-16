using System.IO.Compression;

namespace ExcelReader.Core.Internal
{
    // The "#if NET10_0_OR_GREATER await zip.DisposeAsync() #else zip.Dispose()" idiom, shared by every
    // ZIP-backed reader and writer (XlsxReader, XlsbReader, XlsxWorkbookWriter, XlsbWorkbookWriter)
    // across their dispose paths.
    internal static class ZipArchiveDisposal
    {
        [System.Diagnostics.CodeAnalysis.SuppressMessage("IDisposableAnalyzers.Correctness", "IDISP007:Don't dispose injected",
            Justification = "Disposal helper: disposing the caller-owned ZipArchive is its sole purpose — callers delegate their own zip's disposal here.")]
        internal static ValueTask DisposeAsync(ZipArchive zip)
        {
#if NET10_0_OR_GREATER
            return zip.DisposeAsync();
#else
            zip.Dispose();
            return ValueTask.CompletedTask;
#endif
        }

        // Synchronous counterpart, for native/unmanaged callers whose ABI is synchronous.
        [System.Diagnostics.CodeAnalysis.SuppressMessage("IDisposableAnalyzers.Correctness", "IDISP007:Don't dispose injected",
            Justification = "Disposal helper: disposing the caller-owned ZipArchive is its sole purpose — callers delegate their own zip's disposal here.")]
        internal static void Dispose(ZipArchive zip)
        {
            zip.Dispose();
        }
    }
}
