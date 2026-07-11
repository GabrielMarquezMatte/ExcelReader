using System.Buffers;
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

        [SuppressMessage("SharpSource", "SS066:DisposableFieldIsNotDisposed",
            Justification = "Disposed in Dispose() when _ownsSource; otherwise the caller owns it.")]
        private readonly Stream? _source;
        private readonly bool _ownsSource;
        private readonly int[] _chain;        // physical sector numbers, in order (pooled, oversized)
        private readonly int _chainLength;    // valid entry count in _chain; the pooled array is larger
        private bool _chainReturned;          // guards against a double pool-return on repeated Dispose
        private readonly ReadOnlyMemory<byte> _memory;

        internal int SectorSize { get; }
        internal long Length { get; }

        private WorkbookStream(Stream? source, bool ownsSource, int[] chain, int chainLength, ReadOnlyMemory<byte> memory, int sectorSize, long length)
        {
            _source = source;
            _ownsSource = ownsSource;
            _chain = chain;
            _chainLength = chainLength;
            _memory = memory;
            SectorSize = sectorSize;
            Length = length;
        }

        internal static WorkbookStream Streamed(Stream source, bool ownsSource, int[] chain, int chainLength, int sectorSize, long length)
        {
            return new WorkbookStream(source, ownsSource, chain, chainLength, default, sectorSize, length);
        }

        internal static WorkbookStream InMemory(ReadOnlyMemory<byte> data)
        {
            return new WorkbookStream(null, ownsSource: false, [], 0, data, sectorSize: 1, data.Length);
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

        // Reads contiguous physical sectors starting from chainIndex into dest.
        // Returns the number of sectors loaded.
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
            _source!.Seek(offset, SeekOrigin.Begin);
            _source.ReadExactly(dest[..(count * SectorSize)]);
            return count;
        }

        [SuppressMessage("IDisposableAnalyzers.Correctness", "IDISP007:Don't dispose injected",
            Justification = "Only disposed when _ownsSource — this type took ownership of the source in that case.")]
        public void Dispose()
        {
            if (_ownsSource)
            {
                _source?.Dispose();
            }
            // Streamed mode rents _chain from the pool; in-memory mode holds the shared empty array.
            if (_chainLength > 0 && !_chainReturned)
            {
                _chainReturned = true;
                ArrayPool<int>.Shared.Return(_chain);
            }
        }
    }
}
