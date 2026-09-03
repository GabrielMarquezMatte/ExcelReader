using System.IO.Compression;

namespace ExcelReader.Core.Internal
{
    // The "#if NET10_0_OR_GREATER await zip.DisposeAsync() #else zip.Dispose()" idiom, shared by every
    // ZIP-backed reader and writer (XlsxReader, XlsbReader, XlsxWorkbookWriter, XlsbWorkbookWriter)
    // across their dispose paths.
    internal static class ZipArchiveDisposal
    {
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
        internal static void Dispose(ZipArchive zip)
        {
            zip.Dispose();
        }
    }
}
