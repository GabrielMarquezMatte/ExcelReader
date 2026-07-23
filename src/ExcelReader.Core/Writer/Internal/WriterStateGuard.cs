using System.IO.Compression;

namespace ExcelReader.Core.Writer.Internal
{
    // Shared WriterState guard checks for the three state-tracking workbook writers (XlsxWorkbookWriter,
    // XlsbWorkbookWriter, XlsWorkbookWriter): each StartAsync/AddSheet/EndAsync call site repeats the
    // same "already disposed" / "wrong state" checks, differing only in the writer's type name and
    // the action being guarded.
    internal static class WriterStateGuard
    {
        internal static void ThrowIfEnded(WriterState state, object writer)
        {
            ObjectDisposedException.ThrowIf(state == WriterState.Ended, writer);
        }

        internal static void RequireCreated(WriterState state, string typeName)
        {
            if (state != WriterState.Created)
            {
                throw new InvalidOperationException($"{typeName} has already been started.");
            }
        }

        internal static void RequireStarted(WriterState state, string typeName, string action)
        {
            if (state != WriterState.Started)
            {
                throw new InvalidOperationException($"{typeName} must be started before {action}.");
            }
        }

        // Guards row-writer re-entrancy at the start of a new row — every sheet writer must reject
        // starting row N+1 while row N's writer is still open.
        internal static void RequireNoActiveRowForStart(bool rowActive, string rowWriterTypeName)
        {
            if (rowActive)
            {
                throw new InvalidOperationException($"The previous {rowWriterTypeName} must be disposed before starting a new row.");
            }
        }

        // Guards ending the sheet while its row writer is still open.
        internal static void RequireNoActiveRowForEnd(bool rowActive, string rowWriterTypeName)
        {
            if (rowActive)
            {
                throw new InvalidOperationException($"The active {rowWriterTypeName} must be disposed before ending the sheet.");
            }
        }

        internal static void ValidateSheetName(string name)
        {
            if (name.Length is 0 or > 31)
            {
                throw new ArgumentException("Sheet names must be 1 to 31 characters.", nameof(name));
            }
            if (name.IndexOfAny([':', '\\', '/', '?', '*', '[', ']']) >= 0)
            {
                throw new ArgumentException("Sheet names cannot contain : \\ / ? * [ or ].", nameof(name));
            }
        }
    }

    // The "#if NET10_0_OR_GREATER await zip.DisposeAsync() #else zip.Dispose()" idiom, shared by the
    // two ZIP-backed writers (XlsxWorkbookWriter, XlsbWorkbookWriter) across their EndAsync/DisposeAsync
    // paths.
    internal static class ZipArchiveDisposal
    {
        [System.Diagnostics.CodeAnalysis.SuppressMessage("IDisposableAnalyzers.Correctness", "IDISP007:Don't dispose injected",
            Justification = "Disposal helper: disposing the caller-owned ZipArchive is its sole purpose — the two ZIP-backed writers delegate their own zip's disposal here.")]
        internal static ValueTask DisposeAsync(ZipArchive zip)
        {
#if NET10_0_OR_GREATER
            return zip.DisposeAsync();
#else
            zip.Dispose();
            return ValueTask.CompletedTask;
#endif
        }
    }
}
