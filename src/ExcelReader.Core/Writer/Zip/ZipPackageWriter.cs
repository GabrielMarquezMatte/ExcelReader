using System.Diagnostics.CodeAnalysis;
using System.IO.Compression;
using ExcelReader.Core.Writer.Internal;

namespace ExcelReader.Core.Writer.Zip
{
    /// <summary>
    /// Writes a ZIP archive one entry at a time to a stream it does not own, in the layout
    /// <c>ZipArchive</c> uses in create mode. Offsets are absolute positions in that stream.
    /// </summary>
    [SuppressMessage("Design", "CA1001:Types that own disposable fields should be disposable",
        Justification = "The open entry belongs to the caller that opened it and is completed by its Dispose; Abandon releases it without writing.")]
    internal sealed class ZipPackageWriter
    {
        private readonly Stream _output;
        private readonly long _origin;
        private readonly List<ZipEntryRecord> _entries = [];
        private ZipEntryStream? _open;
        private long _written;
        private bool _closed;
        private bool _broken;

        internal ZipPackageWriter(Stream output)
        {
            _output = output;
            CanSeek = output.CanSeek;
            if (CanSeek)
            {
                _origin = output.Position;
            }
        }

        internal bool CanSeek { get; }

        internal long Position
        {
            get
            {
                return _origin + _written;
            }
        }

        internal Stream OpenEntry(string name, CompressionLevel level)
        {
            ObjectDisposedException.ThrowIf(_closed, this);
            if (!Enum.IsDefined(level))
            {
                throw new ArgumentOutOfRangeException(nameof(level), level, "Unknown compression level.");
            }
            if (_open is not null)
            {
                throw new InvalidOperationException("The previous ZIP entry must be closed before another is opened.");
            }
            if (_broken)
            {
                throw new InvalidOperationException("A ZIP entry failed, so the archive cannot take more entries.");
            }
            _open = new ZipEntryStream(this, new ZipEntryRecord(name, level, Position, DateTime.Now));
            return _open;
        }

        internal void Finish()
        {
            using BiffBuffer directory = BuildCentralDirectory();
            WriteRaw(directory.Span);
        }

        internal async ValueTask FinishAsync(CancellationToken ct = default)
        {
            using BiffBuffer directory = BuildCentralDirectory();
            await WriteRawAsync(directory.Memory, ct).ConfigureAwait(false);
        }

        internal void Abandon()
        {
            _closed = true;
            _open?.Abandon();
            _open = null;
        }

        internal void WriteRaw(ReadOnlySpan<byte> bytes)
        {
            _output.Write(bytes);
            _written += bytes.Length;
        }

        internal async ValueTask WriteRawAsync(ReadOnlyMemory<byte> bytes, CancellationToken ct)
        {
            await _output.WriteAsync(bytes, ct).ConfigureAwait(false);
            _written += bytes.Length;
        }

        internal void Patch(long position, ReadOnlySpan<byte> bytes)
        {
            _output.Seek(position, SeekOrigin.Begin);
            _output.Write(bytes);
            _output.Seek(Position, SeekOrigin.Begin);
        }

        internal async ValueTask PatchAsync(long position, ReadOnlyMemory<byte> bytes, CancellationToken ct)
        {
            _output.Seek(position, SeekOrigin.Begin);
            await _output.WriteAsync(bytes, ct).ConfigureAwait(false);
            _output.Seek(Position, SeekOrigin.Begin);
        }

        internal void EntryClosed(ZipEntryRecord record)
        {
            _entries.Add(record);
            _open = null;
        }

        internal void EntryFailed()
        {
            _open = null;
            _broken = true;
        }

        private BiffBuffer BuildCentralDirectory()
        {
            ObjectDisposedException.ThrowIf(_closed, this);
            if (_open is not null || _broken)
            {
                throw new InvalidOperationException("Every ZIP entry must be closed successfully before the archive is finished.");
            }
            _closed = true;
            long directoryOffset = Position;
            var directory = new BiffBuffer(256 + (_entries.Count * 128));
            foreach (ZipEntryRecord entry in _entries)
            {
                ZipRecords.WriteCentralHeader(directory, entry);
            }
            ZipRecords.WriteEnd(directory, _entries.Count, directoryOffset, directory.Length);
            return directory;
        }
    }
}
