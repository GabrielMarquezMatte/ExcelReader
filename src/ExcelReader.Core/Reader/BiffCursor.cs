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
        private readonly int _maxSectors;
        private byte[]? _sector;     // current sector buffer window (streamed mode only)
        private int _loadedStart = -1;       // start chain index loaded in _sector
        private int _loadedCount;        // number of sectors loaded in _sector
        private byte[]? _scratch;             // assembles records that span sectors

        internal BiffCursor(WorkbookStream wb)
        {
            _wb = wb;
            _sectorSize = wb.SectorSize;
            if (wb.IsMemory)
            {
                _maxSectors = 0;
                _sector = null;
                return;
            }
            _maxSectors = Math.Max(1, 65536 / wb.SectorSize);
            _sector = ArrayPool<byte>.Shared.Rent(_maxSectors * wb.SectorSize);
        }

        internal long Position { get; set; }

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
            int chainIndex = (int)(pos / _sectorSize);
            int within = (int)(pos % _sectorSize);
            LoadSector(chainIndex);
            int offsetInSector = (chainIndex - _loadedStart) * _sectorSize + within;
            if (offsetInSector + len <= _loadedCount * _sectorSize)
            {
                return _sector.AsSpan(offsetInSector, len);
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
                int chainIndex = (int)(at / _sectorSize);
                int within = (int)(at % _sectorSize);
                LoadSector(chainIndex);
                int offsetInSector = (chainIndex - _loadedStart) * _sectorSize + within;
                int availableInSector = _loadedCount * _sectorSize - offsetInSector;
                int take = Math.Min(availableInSector, dest.Length - written);
                _sector.AsSpan(offsetInSector, take).CopyTo(dest[written..]);
                written += take;
            }
        }

        private void LoadSector(int chainIndex)
        {
            if (chainIndex >= _loadedStart && chainIndex < _loadedStart + _loadedCount)
            {
                return;
            }
            int sectorsRead = _wb.LoadSectors(chainIndex, _sector.AsSpan(0, _maxSectors * _sectorSize));
            _loadedStart = chainIndex;
            _loadedCount = sectorsRead;
        }

        private byte[] EnsureScratch(int len)
        {
            if (_scratch is not null && _scratch.Length >= len)
            {
                return _scratch;
            }
            if (_scratch is not null)
            {
                ArrayPool<byte>.Shared.Return(_scratch);
            }
            _scratch = ArrayPool<byte>.Shared.Rent(len);
            return _scratch;
        }

        public void Dispose()
        {
            if (_scratch is not null)
            {
                ArrayPool<byte>.Shared.Return(_scratch);
                _scratch = null;
            }
            if (_sector is not null)
            {
                ArrayPool<byte>.Shared.Return(_sector);
                _sector = null;
            }
        }
    }
}
