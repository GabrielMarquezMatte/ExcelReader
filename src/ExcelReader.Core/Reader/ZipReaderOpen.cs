using System.Diagnostics.CodeAnalysis;
using System.IO.Compression;
using ExcelReader.Core.Writer.Internal;

namespace ExcelReader.Core.Reader
{
    // Shared CreateAsync scaffolding for the ZIP-backed readers (XlsxReader, XlsbReader): open the
    // archive (.NET 10 async API, or the sync ctor as a fallback on earlier targets), run the
    // format-specific part-parsing body, and on any failure dispose the zip and (unless leaveOpen)
    // the stream before rethrowing. parseBody owns zip-entry reads and returns the fully-constructed
    // reader; ownership of `zip` transfers to that returned reader on success.
    internal static class ZipReaderOpen
    {
        [SuppressMessage("IDisposableAnalyzers.Correctness", "IDISP001:Dispose created",
            Justification = "zip ownership transfers to parseBody's returned reader on success; disposed here in the catch on failure.")]
        internal static async ValueTask<TResult> OpenAsync<TResult>(
            Stream stream, bool leaveOpen, CancellationToken ct, Func<ZipArchive, ValueTask<TResult>> parseBody)
        {
            ZipArchive? zip = null;
            try
            {
#if NET10_0_OR_GREATER
                zip = await ZipArchive.CreateAsync(stream, ZipArchiveMode.Read, leaveOpen: true, entryNameEncoding: null, ct).ConfigureAwait(false);
#else
                ct.ThrowIfCancellationRequested();
                zip = new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen: true);
#endif
                return await parseBody(zip).ConfigureAwait(false);
            }
            catch
            {
                if (zip is not null)
                {
                    await ZipArchiveDisposal.DisposeAsync(zip).ConfigureAwait(false);
                }
                if (!leaveOpen)
                {
                    await stream.DisposeAsync().ConfigureAwait(false);
                }
                throw;
            }
        }
    }
}
