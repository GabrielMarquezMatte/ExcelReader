using System.Buffers;
using System.Buffers.Binary;

namespace ExcelReader.Core.Reader
{
    // Forward/seekable record cursor over a WorkbookStream. Each consumer (globals parse, each
    // enumerator) holds its own, so positions never collide; the shared source is repositioned on
    // every sector load. Not thread-safe across cursors used concurrently from multiple threads.
    internal sealed class BiffCursor : IDisposable
    {
        [System.Diagnostics.CodeAnalysis.SuppressMessage("SharpSource", "SS066:DisposableFieldIsNotDisposed",
            Justification = "Borrowed WorkbookStream; its lifetime is owned by XlsReader, not this cursor.")]
        private readonly WorkbookStream _wb;
        private readonly int _sectorSize;
        private readonly byte[]? _sector;     // current sector (streamed mode only)
        private int _loaded = -1;             // chain index in _sector
        private byte[]? _scratch;             // assembles records that span sectors

        internal BiffCursor(WorkbookStream wb)
        {
            _wb = wb;
            _sectorSize = wb.SectorSize;
            _sector = wb.IsMemory ? null : new byte[wb.SectorSize];
        }

        internal long Position { get; set; }

        internal long Length => _wb.Length;

        internal int PeekId()
        {
            if (Position + 4 > _wb.Length)
            {
                return -1;
            }
            ReadOnlySpan<byte> hdr = ReadSpan(Position, 2);
            return BinaryPrimitives.ReadUInt16LittleEndian(hdr);
        }

        internal bool TryReadRecord(out int id, out ReadOnlySpan<byte> data)
        {
            id = 0;
            data = default;
            if (Position + 4 > _wb.Length)
            {
                return false;
            }
            Span<byte> hdr = stackalloc byte[4];
            ReadInto(Position, hdr);
            id = BinaryPrimitives.ReadUInt16LittleEndian(hdr);
            int len = BinaryPrimitives.ReadUInt16LittleEndian(hdr[2..]);
            long dataPos = Position + 4;
            if (dataPos + len > _wb.Length)
            {
                return false;
            }
            data = ReadSpan(dataPos, len);
            Position = dataPos + len;
            return true;
        }

        // A contiguous view of [pos, pos+len). Zero-copy when in-memory or within one sector
        // otherwise assembled into the scratch buffer. Valid only until the next cursor read.
        private ReadOnlySpan<byte> ReadSpan(long pos, int len)
        {
            if (_wb.IsMemory)
            {
                return _wb.Memory(pos, len);
            }
            int within = (int)(pos % _sectorSize);
            if (within + len <= _sectorSize)
            {
                LoadSector((int)(pos / _sectorSize));
                return _sector.AsSpan(within, len);
            }
            byte[] scratch = EnsureScratch(len);
            ReadInto(pos, scratch.AsSpan(0, len));
            return scratch.AsSpan(0, len);
        }

        private void ReadInto(long pos, Span<byte> dest)
        {
            if (_wb.IsMemory)
            {
                _wb.Memory(pos, dest.Length).CopyTo(dest);
                return;
            }
            int written = 0;
            while (written < dest.Length)
            {
                long at = pos + written;
                int within = (int)(at % _sectorSize);
                LoadSector((int)(at / _sectorSize));
                int take = Math.Min(_sectorSize - within, dest.Length - written);
                _sector.AsSpan(within, take).CopyTo(dest[written..]);
                written += take;
            }
        }

        private void LoadSector(int chainIndex)
        {
            if (_loaded == chainIndex)
            {
                return;
            }
            _wb.LoadSector(chainIndex, _sector);
            _loaded = chainIndex;
        }

        private byte[] EnsureScratch(int len)
        {
            if (_scratch is null || _scratch.Length < len)
            {
                if (_scratch is not null)
                {
                    ArrayPool<byte>.Shared.Return(_scratch);
                }
                _scratch = ArrayPool<byte>.Shared.Rent(len);
            }
            return _scratch;
        }

        public void Dispose()
        {
            if (_scratch is not null)
            {
                ArrayPool<byte>.Shared.Return(_scratch);
                _scratch = null;
            }
        }
    }
}
