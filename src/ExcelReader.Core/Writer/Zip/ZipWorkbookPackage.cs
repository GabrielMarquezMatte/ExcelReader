using System.IO.Compression;
using ExcelReader.Core.Writer.Internal;

namespace ExcelReader.Core.Writer.Zip
{
    /// <summary>The workbook-level parts a ZIP-packaged workbook writes once its last sheet has ended.</summary>
    internal interface IZipWorkbookParts
    {
        int SheetCount { get; }

        void WriteParts();

        ValueTask WritePartsAsync(CancellationToken ct);
    }

    /// <summary>A sheet that a <see cref="ZipWorkbookPackage"/> ends or abandons along with its workbook.</summary>
    internal interface IZipSheet : IDisposable, IAsyncDisposable
    {
        void Abandon();

        ValueTask AbandonAsync();
    }

    /// <summary>
    /// The end, abandon and dispose lifecycle shared by the XLSX and XLSB workbook writers: one open
    /// sheet at a time, the workbook parts written at end, and the destination released exactly once.
    /// </summary>
    internal sealed class ZipWorkbookPackage
    {
        private readonly IZipWorkbookParts _parts;
        private readonly string _format;
        private readonly Stream _stream;
        private readonly bool _leaveOpen;
        private readonly CompressionLevel _compression;
        private bool _faulted;
        private bool _zipClosed;
        private bool _disposed;

        internal ZipWorkbookPackage(IZipWorkbookParts parts, string format, Stream stream, bool leaveOpen, CompressionLevel compression)
        {
            _parts = parts;
            _format = format;
            _stream = stream;
            _leaveOpen = leaveOpen;
            _compression = compression;
            Zip = new ZipPackageWriter(stream);
        }

        internal ZipPackageWriter Zip { get; }

        internal bool Ended { get; private set; }

        internal IZipSheet? ActiveSheet { get; set; }

        internal void SheetEnded(bool faulted)
        {
            ActiveSheet = null;
            if (faulted)
            {
                Ended = true;
                _faulted = true;
            }
        }

        internal void End()
        {
            ObjectDisposedException.ThrowIf(Ended, _parts);
            Ended = true;
            try
            {
                ActiveSheet?.Dispose();
                RequireSheet();
                _parts.WriteParts();
                _zipClosed = true;
                Zip.Finish();
            }
            catch
            {
                Abandon();
                throw;
            }
        }

        internal async ValueTask EndAsync(CancellationToken ct)
        {
            ObjectDisposedException.ThrowIf(Ended, _parts);
            ct.ThrowIfCancellationRequested();
            Ended = true;
            try
            {
                if (ActiveSheet is not null)
                {
                    await ActiveSheet.DisposeAsync().ConfigureAwait(false);
                }
                RequireSheet();
                await _parts.WritePartsAsync(ct).ConfigureAwait(false);
                _zipClosed = true;
                await Zip.FinishAsync(ct).ConfigureAwait(false);
            }
            catch
            {
                await AbandonAsync().ConfigureAwait(false);
                throw;
            }
        }

        private void RequireSheet()
        {
            if (_parts.SheetCount == 0)
            {
                throw new InvalidOperationException($"An {_format} workbook must contain at least one sheet.");
            }
        }

        private void Abandon()
        {
            _faulted = true;
            ActiveSheet?.Abandon();
            if (!_zipClosed)
            {
                _zipClosed = true;
                Zip.Abandon();
            }
        }

        private async ValueTask AbandonAsync()
        {
            _faulted = true;
            if (ActiveSheet is not null)
            {
                await ActiveSheet.AbandonAsync().ConfigureAwait(false);
            }
            if (!_zipClosed)
            {
                _zipClosed = true;
                await Zip.AbandonAsync().ConfigureAwait(false);
            }
        }

        internal void Flush()
        {
            _stream.Flush();
        }

        internal ValueTask FlushAsync(CancellationToken ct)
        {
            return ZipEntryWriter.FlushAsync(_stream, ct);
        }

        internal void Dispose()
        {
            if (_disposed)
            {
                return;
            }
            _disposed = true;
            try
            {
                Close();
            }
            catch
            {
                _faulted = true;
                ReleaseStream();
                throw;
            }
            ReleaseStream();
        }

        internal async ValueTask DisposeAsync()
        {
            if (_disposed)
            {
                return;
            }
            _disposed = true;
            try
            {
                await CloseAsync().ConfigureAwait(false);
            }
            catch
            {
                _faulted = true;
                await ReleaseStreamAsync().ConfigureAwait(false);
                throw;
            }
            await ReleaseStreamAsync().ConfigureAwait(false);
        }

        private void Close()
        {
            if (Ended)
            {
                if (!_zipClosed)
                {
                    Abandon();
                }
                return;
            }
            if (_parts.SheetCount == 0 && ActiveSheet is null)
            {
                Ended = true;
                _zipClosed = true;
                Zip.Finish();
                return;
            }
            End();
        }

        private async ValueTask CloseAsync()
        {
            if (Ended)
            {
                if (!_zipClosed)
                {
                    await AbandonAsync().ConfigureAwait(false);
                }
                return;
            }
            if (_parts.SheetCount == 0 && ActiveSheet is null)
            {
                Ended = true;
                _zipClosed = true;
                await Zip.FinishAsync().ConfigureAwait(false);
                return;
            }
            await EndAsync(CancellationToken.None).ConfigureAwait(false);
        }

        private void ReleaseStream()
        {
            if (_leaveOpen)
            {
                return;
            }
            if (_faulted)
            {
                FailureCleanup.Dispose(_stream);
                return;
            }
            _stream.Dispose();
        }

        private ValueTask ReleaseStreamAsync()
        {
            if (_leaveOpen)
            {
                return ValueTask.CompletedTask;
            }
            return _faulted ? FailureCleanup.DisposeAsync(_stream) : _stream.DisposeAsync();
        }

        internal void WriteEntry(string entryName, string content)
        {
            ZipEntryWriter.WriteText(Zip, entryName, content, _compression);
        }

        internal void WriteEntry(string entryName, ReadOnlySpan<byte> content)
        {
            ZipEntryWriter.WriteBytes(Zip, entryName, content, _compression);
        }

        internal ValueTask WriteEntryAsync(string entryName, string content, CancellationToken ct)
        {
            return ZipEntryWriter.WriteTextAsync(Zip, entryName, content, _compression, ct);
        }

        internal ValueTask WriteEntryAsync(string entryName, ReadOnlyMemory<byte> content, CancellationToken ct)
        {
            return ZipEntryWriter.WriteBytesAsync(Zip, entryName, content, _compression, ct);
        }

        internal void WriteSharedStrings(string entryName, SharedStringTable strings, bool binary)
        {
            using Stream entry = Zip.OpenEntry(entryName, _compression);
            strings.WriteTo(entry, binary);
        }

        internal async ValueTask WriteSharedStringsAsync(string entryName, SharedStringTable strings, bool binary, CancellationToken ct)
        {
            Stream entry = Zip.OpenEntry(entryName, _compression);
            await using (entry.ConfigureAwait(false))
            {
                await strings.WriteToAsync(entry, binary, ct).ConfigureAwait(false);
            }
        }
    }
}
