using System.Buffers;
using System.Buffers.Binary;
using System.Runtime.CompilerServices;

namespace ExcelReader.Core.Reader.Xls
{
    internal sealed class BiffCursor : IDisposable
    {
        private readonly WorkbookStream _wb;
        private readonly WorkbookStream.SourceKind _kind;
        private readonly int _sectorSize;
        private readonly int _maxSectors;
        private byte[]? _sector;
        private int _loadedStart = -1;
        private int _loadedCount;
        private byte[]? _scratch;

        private readonly byte[] _file;
        private readonly int _fileBase;

        private long _runStart = -1;
        private long _runEnd = -1;
        private long _runBufferOffset;

        private long _windowStart = -1;
        private long _windowEnd = -1;

        internal BiffCursor(WorkbookStream wb)
        {
            _wb = wb;
            _kind = wb.Kind;
            Length = wb.Length;
            _sectorSize = wb.SectorSize;
            _file = wb.Buffer;
            _fileBase = wb.BufferBase;
            if (wb.Kind != WorkbookStream.SourceKind.Streamed)
            {
                _maxSectors = 0;
                _sector = null;
                return;
            }
            _maxSectors = Math.Max(1, 65536 / wb.SectorSize);
            _sector = ArrayPool<byte>.Shared.Rent(_maxSectors * wb.SectorSize);
        }

        internal long Position { get; set; }

        internal long Length { get; }

        internal int PeekId()
        {
            if (Position + 4 > Length)
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
            if (Position + 4 > Length)
            {
                return false;
            }
            ReadOnlySpan<byte> hdr = ReadSpan(Position, 4);
            id = BinaryPrimitives.ReadUInt16LittleEndian(hdr);
            int len = BinaryPrimitives.ReadUInt16LittleEndian(hdr[2..]);
            long dataPos = Position + 4;
            if (dataPos + len > Length)
            {
                return false;
            }
            data = ReadSpan(dataPos, len);
            Position = dataPos + len;
            return true;
        }

        private ReadOnlySpan<byte> ReadSpan(long pos, int len)
        {
            if (_kind == WorkbookStream.SourceKind.Contiguous)
            {
                return _wb.Memory(pos, len);
            }
            if (_kind == WorkbookStream.SourceKind.Chained)
            {
                return ReadChainedSpan(pos, len);
            }
            if (pos >= _windowStart && pos + len <= _windowEnd)
            {
                return _sector.AsSpan((int)(pos - _windowStart), len);
            }
            return ReadStreamedSpanSlow(pos, len);
        }

        [MethodImpl(MethodImplOptions.NoInlining)]
        private ReadOnlySpan<byte> ReadStreamedSpanSlow(long pos, int len)
        {
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

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        private ReadOnlySpan<byte> ReadChainedSpan(long pos, int len)
        {
            if (pos >= _runStart && pos + len <= _runEnd)
            {
                return _file.AsSpan(_fileBase + (int)(_runBufferOffset + pos - _runStart), len);
            }
            return ReadChainedSpanSlow(pos, len);
        }

        private ReadOnlySpan<byte> ReadChainedSpanSlow(long pos, int len)
        {
            _wb.ResolveChainedRun(pos, out _runStart, out _runEnd, out _runBufferOffset);
            if (pos + len <= _runEnd)
            {
                return _file.AsSpan(_fileBase + (int)(_runBufferOffset + pos - _runStart), len);
            }
            byte[] scratch = EnsureScratch(len);
            _wb.CopyChained(pos, scratch.AsSpan(0, len));
            return scratch.AsSpan(0, len);
        }

        private void ReadInto(long pos, Span<byte> dest)
        {
            if (_kind == WorkbookStream.SourceKind.Contiguous)
            {
                _wb.Memory(pos, dest.Length).CopyTo(dest);
                return;
            }
            if (_kind == WorkbookStream.SourceKind.Chained)
            {
                _wb.CopyChained(pos, dest);
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
            _windowStart = (long)chainIndex * _sectorSize;
            _windowEnd = _windowStart + ((long)sectorsRead * _sectorSize);
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
