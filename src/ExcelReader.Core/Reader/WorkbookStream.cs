using System.Buffers;
using System.Buffers.Binary;
using System.Diagnostics.CodeAnalysis;

namespace ExcelReader.Core.Reader
{
    // The Workbook OLE stream, read on demand instead of materialized. Two modes:
    //  - streamed: a seekable source + the stream's physical FAT-sector chain. Only one sector
    //    (plus a record-assembly scratch) is held at a time, so a 3 MB workbook costs ~KBs.
    //  - in-memory: a contiguous buffer (mini-stream workbooks, or a non-seekable fallback).
    // Immutable and shareable; each consumer reads through its own BiffCursor.
    [ExcludeFromCodeCoverage(Justification = "Exercised through XlsReader integration tests; guard-rail branches are corrupt-OLE only.")]
    internal sealed class WorkbookStream : IDisposable
    {
        private const int HeaderSize = 512;

        private readonly Stream? _source;
        private readonly bool _ownsSource;
        private readonly int[] _chain;        // physical sector numbers, in order
        private readonly ReadOnlyMemory<byte> _memory;

        internal int SectorSize { get; }
        internal long Length { get; }

        private WorkbookStream(Stream? source, bool ownsSource, int[] chain, ReadOnlyMemory<byte> memory, int sectorSize, long length)
        {
            _source = source;
            _ownsSource = ownsSource;
            _chain = chain;
            _memory = memory;
            SectorSize = sectorSize;
            Length = length;
        }

        internal static WorkbookStream Streamed(Stream source, bool ownsSource, int[] chain, int sectorSize, long length)
        {
            return new WorkbookStream(source, ownsSource, chain, default, sectorSize, length);
        }

        internal static WorkbookStream InMemory(ReadOnlyMemory<byte> data)
        {
            return new WorkbookStream(null, ownsSource: false, [], data, sectorSize: 1, data.Length);
        }

        internal BiffCursor OpenCursor()
        {
            return new BiffCursor(this);
        }

        internal bool IsMemory => _source is null;

        internal ReadOnlySpan<byte> Memory(long pos, int len)
        {
            return _memory.Span.Slice((int)pos, len);
        }

        // Reads sectorSize bytes of physical sector at chainIndex into dest (length == SectorSize).

        internal void LoadSector(int chainIndex, Span<byte> dest)
        {
            if ((uint)chainIndex >= (uint)_chain.Length)
            {
                throw new InvalidDataException("Invalid OLE sector chain index.");
            }
            long offset = HeaderSize + ((long)_chain[chainIndex] * SectorSize);
            _source!.Seek(offset, SeekOrigin.Begin);
            _source.ReadExactly(dest);
        }

        public void Dispose()
        {
            if (_ownsSource)
            {
                _source?.Dispose();
            }
        }
    }

    // Forward/seekable record cursor over a WorkbookStream. Each consumer (globals parse, each
    // enumerator) holds its own, so positions never collide; the shared source is repositioned on
    // every sector load. Not thread-safe across cursors used concurrently from multiple threads.
    internal sealed class BiffCursor : IDisposable
    {
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
