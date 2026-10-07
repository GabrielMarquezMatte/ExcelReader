using System.Buffers;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;
using ExcelReader.Core.Reader.Sources;

namespace ExcelReader.Core.Reader.Xls
{
    [ExcludeFromCodeCoverage(Justification = "Exercised through XlsWorkbook integration tests; guard-rail branches are corrupt-OLE only.")]
    internal sealed class WorkbookStream : IDisposable
    {
        private const int HeaderSize = 512;

        private readonly ByteSource? _source;
        private readonly int[] _chain;
        private readonly int _chainLength;
        private bool _chainReturned;

        private readonly int _fileLength;

        internal int SectorSize { get; }
        internal long Length { get; }

        internal enum SourceKind
        {
            Streamed,
            Contiguous,
            Chained,
        }

        internal SourceKind Kind { get; }

        private WorkbookStream(ByteSource? source, int[] chain, int chainLength, ReadOnlyMemory<byte> memory, int sectorSize, long length, SourceKind kind)
        {
            _source = source;
            _chain = chain;
            _chainLength = chainLength;
            SectorSize = sectorSize;
            Length = length;
            Kind = kind;
            if (kind == SourceKind.Streamed)
            {
                Buffer = [];
                return;
            }
            if (MemoryMarshal.TryGetArray(memory, out ArraySegment<byte> segment))
            {
                Buffer = segment.Array!;
                BufferBase = segment.Offset;
                _fileLength = segment.Count;
                return;
            }
            Buffer = memory.ToArray();
            _fileLength = Buffer.Length;
        }

        internal static WorkbookStream Streamed(ByteSource source, int[] chain, int chainLength, int sectorSize, long length)
        {
            return new WorkbookStream(source, chain, chainLength, default, sectorSize, length, SourceKind.Streamed);
        }

        internal static WorkbookStream InMemory(ReadOnlyMemory<byte> data)
        {
            return new WorkbookStream(null, [], 0, data, sectorSize: 1, data.Length, SourceKind.Contiguous);
        }

        internal static WorkbookStream Chained(ReadOnlyMemory<byte> data, int[] chain, int chainLength, int sectorSize, long length)
        {
            return new WorkbookStream(null, chain, chainLength, data, sectorSize, length, SourceKind.Chained);
        }

        internal BiffCursor OpenCursor()
        {
            return new BiffCursor(this);
        }


        internal byte[] Buffer { get; }
        internal int BufferBase { get; }

        internal ReadOnlySpan<byte> Memory(long pos, int len)
        {
            return Buffer.AsSpan(BufferBase + (int)pos, len);
        }

        internal void ResolveChainedRun(long pos, out long runStart, out long runEnd, out long bufferOffset)
        {
            int chainIndex = (int)(pos / SectorSize);
            RequireChainIndex(chainIndex);
            int first = chainIndex;
            while (first > 0 && _chain[first] == _chain[first - 1] + 1)
            {
                first--;
            }
            int last = chainIndex;
            while (last + 1 < _chainLength && _chain[last + 1] == _chain[last] + 1)
            {
                last++;
            }
            runStart = (long)first * SectorSize;
            runEnd = Math.Min(Length, (long)(last + 1) * SectorSize);
            bufferOffset = HeaderSize + ((long)_chain[first] * SectorSize);
            if (_chain[first] < 0 || bufferOffset < 0 || bufferOffset + (runEnd - runStart) > _fileLength)
            {
                throw new InvalidDataException("The OLE sector chain points past the end of the buffer.");
            }
        }

        internal void CopyChained(long pos, Span<byte> dest)
        {
            int written = 0;
            while (written < dest.Length)
            {
                long at = pos + written;
                int chainIndex = (int)(at / SectorSize);
                int within = (int)(at % SectorSize);
                RequireChainIndex(chainIndex);
                int take = Math.Min(SectorSize - within, dest.Length - written);
                FileSlice(_chain[chainIndex], within, take).CopyTo(dest[written..]);
                written += take;
            }
        }

        private void RequireChainIndex(int chainIndex)
        {
            if ((uint)chainIndex >= (uint)_chainLength)
            {
                throw new InvalidDataException("Invalid OLE sector chain index.");
            }
        }

        private ReadOnlySpan<byte> FileSlice(int sector, int within, int len)
        {
            long offset = HeaderSize + ((long)sector * SectorSize) + within;
            if (sector < 0 || offset < 0 || offset + len > _fileLength)
            {
                throw new InvalidDataException("The OLE sector chain points past the end of the buffer.");
            }
            return Buffer.AsSpan(BufferBase + (int)offset, len);
        }

        internal int LoadSectors(int chainIndex, Span<byte> dest)
        {
            if ((uint)chainIndex >= (uint)_chainLength)
            {
                throw new InvalidDataException("Invalid OLE sector chain index.");
            }
            int maxSectors = dest.Length / SectorSize;
            int count = 1;
            while (count < maxSectors && chainIndex + count < _chainLength)
            {
                if (_chain[chainIndex + count] != _chain[chainIndex + count - 1] + 1)
                {
                    break;
                }
                count++;
            }
            long offset = HeaderSize + ((long)_chain[chainIndex] * SectorSize);
            _source!.ReadExactly(offset, dest[..(count * SectorSize)]);
            return count;
        }

        public void Dispose()
        {
            _source?.Dispose();
            if (_chainLength > 0 && !_chainReturned)
            {
                _chainReturned = true;
                ArrayPool<int>.Shared.Return(_chain);
            }
        }
    }
}
