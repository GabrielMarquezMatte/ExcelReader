using System.Buffers;
using System.Buffers.Binary;
using System.Numerics;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using ExcelReader.Core.Writer.Internal;

namespace ExcelReader.Core.Writer.Zip
{
    /// <summary>
    /// Raw deflate (RFC 1951) for <c>CompressionLevel.Fastest</c>: one hash candidate per position, a
    /// 4-byte minimum match, no lazy matching and fixed Huffman codes. Input is encoded in blocks of at
    /// most <see cref="MaxBlockLength"/> bytes, and a block the fixed codes would enlarge is written
    /// stored instead.
    /// </summary>
    internal sealed class DeflateEncoder : IDisposable
    {
        internal const int MaxBlockLength = 65_535;

        private const int HashBits = 16;
        private const int HistoryLength = 32_768;
        private const int MinMatch = 4;
        private const int MaxMatch = 258;
        private const int WindowCapacity = HistoryLength + (1 << 20);
        // Every byte a 9-bit literal, plus the block header, end-of-block, pending bits and a 4-byte flush.
        private const int MaxEncodedBlock = (MaxBlockLength * 9 / 8) + 16;

        private byte[] _window;
        private uint[] _table;
        private int _encoded;
        private int _end;
        private uint _windowStart;
        private ulong _bits;
        private int _bitCount;

        internal DeflateEncoder()
        {
            _window = ArrayPool<byte>.Shared.Rent(WindowCapacity);
            _table = ArrayPool<uint>.Shared.Rent(1 << HashBits);
            Array.Clear(_table, 0, 1 << HashBits);
        }

        internal void Write(ReadOnlySpan<byte> input, BiffBuffer output)
        {
            ObjectDisposedException.ThrowIf(_window.Length == 0, this);
            while (!input.IsEmpty)
            {
                if (_end == WindowCapacity)
                {
                    Slide();
                }
                int take = Math.Min(input.Length, WindowCapacity - _end);
                input[..take].CopyTo(_window.AsSpan(_end));
                _end += take;
                input = input[take..];
                while (_end - _encoded >= MaxBlockLength)
                {
                    EncodeBlock(_encoded + MaxBlockLength, final: false, output);
                }
            }
        }

        internal void Finish(BiffBuffer output)
        {
            EncodeBlock(_end, final: true, output);
            Span<byte> tail = output.GetSpan(sizeof(ulong));
            BinaryPrimitives.WriteUInt64LittleEndian(tail, _bits);
            output.Advance((_bitCount + 7) / 8);
            _bits = 0;
            _bitCount = 0;
        }

        public void Dispose()
        {
            if (_window.Length == 0)
            {
                return;
            }
            ArrayPool<byte>.Shared.Return(_window);
            ArrayPool<uint>.Shared.Return(_table);
            _window = [];
            _table = [];
        }

        // Keeps the last 32 KB already encoded as history. The hash table holds absolute positions, so
        // moving the window only moves _windowStart; it never rewrites the table.
        private void Slide()
        {
            int keepFrom = Math.Max(0, _encoded - HistoryLength);
            _window.AsSpan(keepFrom, _end - keepFrom).CopyTo(_window);
            _windowStart += (uint)keepFrom;
            _encoded -= keepFrom;
            _end -= keepFrom;
        }

        private void EncodeBlock(int blockEnd, bool final, BiffBuffer output)
        {
            Span<byte> destination = output.GetSpan(MaxEncodedBlock);
            int length = blockEnd - _encoded;
            var state = new BitState
            {
                Bits = _bits | ((final ? 3UL : 2UL) << _bitCount),
                Count = _bitCount + 3,
            };
            // Up to 31 pending bits plus the 3 header bits can exceed 32; flush so a 31-bit symbol still fits.
            FlushWord(destination, ref state);
            int next = EncodeMatches(blockEnd, destination, ref state);
            EncodeLiterals(next, blockEnd, destination, ref state);
            state.Count += 7;
            FlushWord(destination, ref state);
            long fixedBits = (8L * state.Written) + state.Count - _bitCount;
            if (fixedBits <= StoredBits(length))
            {
                output.Advance(state.Written);
                _bits = state.Bits;
                _bitCount = state.Count;
            }
            else
            {
                output.Advance(WriteStored(destination, length, final));
            }
            _encoded = blockEnd;
        }

        private long StoredBits(int length)
        {
            int padding = (8 - ((_bitCount + 3) & 7)) & 7;
            return 3 + padding + 32 + (8L * length);
        }

        private int WriteStored(Span<byte> destination, int length, bool final)
        {
            ulong header = _bits | ((final ? 1UL : 0UL) << _bitCount);
            int headerBytes = (_bitCount + 3 + 7) / 8;
            BinaryPrimitives.WriteUInt64LittleEndian(destination, header);
            BinaryPrimitives.WriteUInt16LittleEndian(destination[headerBytes..], (ushort)length);
            BinaryPrimitives.WriteUInt16LittleEndian(destination[(headerBytes + 2)..], (ushort)~length);
            _window.AsSpan(_encoded, length).CopyTo(destination[(headerBytes + 4)..]);
            _bits = 0;
            _bitCount = 0;
            return headerBytes + 4 + length;
        }

        private int EncodeMatches(int blockEnd, Span<byte> destination, ref BitState state)
        {
            byte[] windowArray = _window;
            uint[] tableArray = _table;
            ObjectDisposedException.ThrowIf(windowArray.Length < WindowCapacity || tableArray.Length < 1 << HashBits, this);
            ref byte window = ref MemoryMarshal.GetArrayDataReference(windowArray);
            ref uint table = ref MemoryMarshal.GetArrayDataReference(tableArray);
            ref byte output = ref MemoryMarshal.GetReference(destination);
            ref uint literals = ref MemoryMarshal.GetArrayDataReference(FixedHuffmanCodes.Literals);
            ref uint lengths = ref MemoryMarshal.GetArrayDataReference(FixedHuffmanCodes.Lengths);
            ref uint distances = ref MemoryMarshal.GetArrayDataReference(FixedHuffmanCodes.Distances);
            ulong bits = state.Bits;
            int count = state.Count;
            int written = state.Written;
            int index = _encoded;
            int lastMatchStart = blockEnd - MinMatch;
            while (index <= lastMatchStart)
            {
                uint current = Unsafe.ReadUnaligned<uint>(ref Unsafe.Add(ref window, index));
                ref uint slot = ref Unsafe.Add(ref table, (int)((current * 2654435761u) >> (32 - HashBits)));
                uint position = _windowStart + (uint)index;
                uint distance = position - slot;
                slot = position;
                if (distance - 1 < HistoryLength && distance <= (uint)index
                    && Unsafe.ReadUnaligned<uint>(ref Unsafe.Add(ref window, index - (int)distance)) == current)
                {
                    int length = MatchLength(ref window, index - (int)distance, index, Math.Min(MaxMatch, blockEnd - index));
                    uint lengthEntry = Unsafe.Add(ref lengths, length);
                    uint distanceEntry = Unsafe.Add(ref distances, (int)distance);
                    int lengthBits = (int)(lengthEntry & FixedHuffmanCodes.LengthMask);
                    ulong code = (lengthEntry >> FixedHuffmanCodes.LengthShift)
                        | ((ulong)(distanceEntry >> FixedHuffmanCodes.LengthShift) << lengthBits);
                    bits |= code << count;
                    count += lengthBits + (int)(distanceEntry & FixedHuffmanCodes.LengthMask);
                    index += length;
                }
                else
                {
                    uint literal = Unsafe.Add(ref literals, Unsafe.Add(ref window, index));
                    bits |= (ulong)(literal >> FixedHuffmanCodes.LengthShift) << count;
                    count += (int)(literal & FixedHuffmanCodes.LengthMask);
                    index++;
                }
                if (count >= 32)
                {
                    Unsafe.WriteUnaligned(ref Unsafe.Add(ref output, written), (uint)bits);
                    written += 4;
                    bits >>= 32;
                    count -= 32;
                }
            }
            state.Bits = bits;
            state.Count = count;
            state.Written = written;
            return index;
        }

        private void EncodeLiterals(int index, int blockEnd, Span<byte> destination, ref BitState state)
        {
            for (; index < blockEnd; index++)
            {
                uint literal = FixedHuffmanCodes.Literals[_window[index]];
                state.Bits |= (ulong)(literal >> FixedHuffmanCodes.LengthShift) << state.Count;
                state.Count += (int)(literal & FixedHuffmanCodes.LengthMask);
                FlushWord(destination, ref state);
            }
        }

        private static void FlushWord(Span<byte> destination, ref BitState state)
        {
            if (state.Count < 32)
            {
                return;
            }
            BinaryPrimitives.WriteUInt32LittleEndian(destination[state.Written..], (uint)state.Bits);
            state.Written += 4;
            state.Bits >>= 32;
            state.Count -= 32;
        }

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        private static int MatchLength(ref byte window, int candidate, int current, int maxLength)
        {
            int length = MinMatch;
            while (length + 8 <= maxLength)
            {
                ulong difference = Unsafe.ReadUnaligned<ulong>(ref Unsafe.Add(ref window, candidate + length))
                    ^ Unsafe.ReadUnaligned<ulong>(ref Unsafe.Add(ref window, current + length));
                if (difference != 0)
                {
                    return length + (BitOperations.TrailingZeroCount(difference) >> 3);
                }
                length += 8;
            }
            while (length < maxLength && Unsafe.Add(ref window, candidate + length) == Unsafe.Add(ref window, current + length))
            {
                length++;
            }
            return length;
        }

        [StructLayout(LayoutKind.Auto)]
        private struct BitState
        {
            internal ulong Bits;
            internal int Count;
            internal int Written;
        }
    }
}
