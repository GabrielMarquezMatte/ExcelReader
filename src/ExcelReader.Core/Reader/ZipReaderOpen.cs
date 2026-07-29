using System.Diagnostics.CodeAnalysis;
using System.IO.Compression;
using ExcelReader.Core.Internal;

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
        [SuppressMessage("Reliability", "CA2000:Dispose objects before losing scope",
            Justification = "zip ownership transfers to parseBody's returned reader on success; disposed here in the catch on failure.")]
        internal static async ValueTask<TResult> OpenAsync<TResult>(
            Stream stream, bool leaveOpen, ExcelReaderOptions options, Func<ZipArchive, ValueTask<TResult>> parseBody, CancellationToken ct = default)
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
                LimitChecks.ThrowIfTooManyEntries(zip.Entries.Count, options);
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

        // Async open over an already-opened ZipArchive: the twin of OpenAsync for callers (Excel.OpenAsync's
        // DetectSeekableAsync) that already opened the archive for format detection, so its central directory
        // is not parsed a second time. Bypasses OpenAsync's own archive creation, so dispose-on-failure lives
        // here instead. `parseBody` owns the entry reads and returns the fully-constructed reader; ownership
        // of `zip` transfers to that reader on success.
        [SuppressMessage("IDisposableAnalyzers.Correctness", "IDISP007:Don't dispose injected",
            Justification = "zip and stream ownership transfers to parseBody's returned reader on success; disposed here in the catch on failure.")]
        internal static async ValueTask<TResult> FromOpenZipAsync<TResult>(
            Stream stream, bool leaveOpen, ZipArchive zip, ExcelReaderOptions options,
            Func<ZipArchive, ValueTask<TResult>> parseBody)
        {
            try
            {
                LimitChecks.ThrowIfTooManyEntries(zip.Entries.Count, options);
                return await parseBody(zip).ConfigureAwait(false);
            }
            catch
            {
                await ZipArchiveDisposal.DisposeAsync(zip).ConfigureAwait(false);
                if (!leaveOpen)
                {
                    await stream.DisposeAsync().ConfigureAwait(false);
                }
                throw;
            }
        }

        // Memory-path twin of OpenAsync's dispose-on-failure contract: on success `memZip`'s lifetime
        // transfers to the reader `build` returns; on any failure it is disposed here before rethrowing.
        [SuppressMessage("IDisposableAnalyzers.Correctness", "IDISP007:Don't dispose injected",
            Justification = "memZip's lifetime transfers to this call; disposing it here on failure is correct ownership, not disposing a borrowed dependency.")]
        internal static TResult FromMemory<TResult>(ZipMemoryIndex memZip, Func<ZipMemoryIndex, TResult> build)
        {
            try
            {
                return build(memZip);
            }
            catch
            {
                memZip.Dispose();
                throw;
            }
        }
    }
}
