using System.IO.Compression;
using ExcelReader.Core.Reader.Internal;

namespace ExcelReader.Core.Reader.Zip
{
    internal static class ZipReaderOpen
    {
        internal static ZipArchive Open(Stream stream, bool leaveOpen)
        {
            try
            {
                return new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen: true);
            }
            catch
            {
                if (!leaveOpen)
                {
                    stream.Dispose();
                }
                throw;
            }
        }

        internal static async ValueTask<TResult> OpenAsync<TResult>(
            Stream stream, bool leaveOpen, ExcelReaderOptions options, Func<ZipArchive, ValueTask<TResult>> parseBody, CancellationToken ct = default)
        {
            ZipArchive? zip = null;
            try
            {
                zip = await ZipArchive.CreateAsync(stream, ZipArchiveMode.Read, leaveOpen: true, entryNameEncoding: null, ct).ConfigureAwait(false);
                LimitChecks.ThrowIfTooManyEntries(zip.Entries.Count, options);
                return await parseBody(zip).ConfigureAwait(false);
            }
            catch
            {
                if (zip is not null)
                {
                    await zip.DisposeAsync().ConfigureAwait(false);
                }
                if (!leaveOpen)
                {
                    await stream.DisposeAsync().ConfigureAwait(false);
                }
                throw;
            }
        }

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
                await zip.DisposeAsync().ConfigureAwait(false);
                if (!leaveOpen)
                {
                    await stream.DisposeAsync().ConfigureAwait(false);
                }
                throw;
            }
        }

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
