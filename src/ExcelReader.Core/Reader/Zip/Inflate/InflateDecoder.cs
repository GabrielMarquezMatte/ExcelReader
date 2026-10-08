using System.Buffers;
using System.Buffers.Binary;

namespace ExcelReader.Core.Reader.Zip.Inflate
{
    internal enum InflateStatus
    {
        NeedInput,
        OutputFull,
        Finished,
    }

    /// <summary>
    /// Resumable DEFLATE decoder with no I/O of its own. The caller owns the input buffer and the
    /// output window and feeds them through <see cref="Decode"/>.
    /// </summary>
    internal sealed partial class InflateDecoder : IDisposable
    {
        // Bytes the caller must keep readable past the end of its input, zeroed once input has ended.
        internal const int InputPadding = 2048;

        // Bytes the caller must keep writable past outLimit: one loop iteration writes at most
        // three literals plus a 258-byte match rounded up by its word-at-a-time copy.
        internal const int OutputSlack = 320;

        // A dynamic block header is at most 570 bytes; the decoder never starts one with less than this.
        private const int HeaderMargin = 1024;
        private const int LoopInputMargin = 64;
        private const int LensSize = InflateFormat.MaxLitlenSymbols + InflateFormat.MaxOffsetSymbols + InflateFormat.MaxRepeat;
        private const ulong PrecodeMask = (1UL << InflateFormat.PrecodeTableBits) - 1;

        private readonly uint[] _litlen = ArrayPool<uint>.Shared.Rent(InflateFormat.LitlenTableSize);
        private readonly uint[] _offset = ArrayPool<uint>.Shared.Rent(InflateFormat.OffsetTableSize);
        private ulong _litlenMask;
        private bool _staticCodesLoaded;
        private ulong _bitbuf;
        private int _bitsleft;
        private BlockState _state;
        private bool _finalBlock;
        private int _storedRemaining;
        private bool _disposed;

        private enum BlockState
        {
            Header,
            Coded,
            Stored,
            Finished,
        }

        /// <summary>
        /// Decodes until the window reaches <paramref name="outLimit"/>, more input is required, or the
        /// stream ends. On <see cref="InflateStatus.NeedInput"/> every unconsumed whole byte has been
        /// handed back through <paramref name="inPos"/>, so the caller may compact and refill.
        /// </summary>
        internal InflateStatus Decode(byte[] input, ref int inPos, int inEnd, bool inputEnded, byte[] window, ref int outPos, int outLimit)
        {
            while (_state != BlockState.Finished)
            {
                if (outPos >= outLimit)
                {
                    return InflateStatus.OutputFull;
                }
                bool advanced = _state switch
                {
                    BlockState.Header => TryReadHeader(input, ref inPos, inEnd, inputEnded),
                    BlockState.Stored => TryCopyStored(input, ref inPos, inEnd, inputEnded, window, ref outPos, outLimit),
                    _ => TryDecodeCoded(input, ref inPos, inEnd, inputEnded, window, ref outPos, outLimit),
                };
                if (!advanced)
                {
                    return InflateStatus.NeedInput;
                }
            }
            return InflateStatus.Finished;
        }

        public void Dispose()
        {
            if (_disposed)
            {
                return;
            }
            _disposed = true;
            ArrayPool<uint>.Shared.Return(_litlen);
            ArrayPool<uint>.Shared.Return(_offset);
        }

        private bool TryReadHeader(byte[] input, ref int inPos, int inEnd, bool inputEnded)
        {
            if (!inputEnded && inEnd - inPos < HeaderMargin)
            {
                Unread(ref inPos);
                return false;
            }
            ThrowIfOverrun(inPos, inEnd, inputEnded);
            RefillBits(input, ref inPos);
            _finalBlock = PopBits(1) != 0;
            switch (PopBits(2))
            {
                case 0:
                    BeginStored(input, ref inPos);
                    break;
                case 1:
                    LoadStaticCodes();
                    _state = BlockState.Coded;
                    break;
                case 2:
                    ReadDynamicCodes(input, ref inPos);
                    _state = BlockState.Coded;
                    break;
                default:
                    throw new InvalidDataException("The deflate stream uses a reserved block type.");
            }
            ThrowIfOverrun(inPos, inEnd, inputEnded);
            return true;
        }

        private void BeginStored(ReadOnlySpan<byte> input, ref int inPos)
        {
            inPos -= _bitsleft >> 3;
            _bitbuf = 0;
            _bitsleft = 0;
            ushort length = BinaryPrimitives.ReadUInt16LittleEndian(input[inPos..]);
            ushort complement = BinaryPrimitives.ReadUInt16LittleEndian(input[(inPos + 2)..]);
            inPos += 4;
            if (length != (ushort)~complement)
            {
                throw new InvalidDataException("A stored deflate block has mismatched length fields.");
            }
            _storedRemaining = length;
            _state = BlockState.Stored;
        }

        private bool TryCopyStored(byte[] input, ref int inPos, int inEnd, bool inputEnded, byte[] window, ref int outPos, int outLimit)
        {
            if (_storedRemaining == 0)
            {
                EndBlock();
                return true;
            }
            int available = inEnd - inPos;
            if (available <= 0)
            {
                ThrowIfTruncated(inputEnded);
                return false;
            }
            int count = Math.Min(Math.Min(_storedRemaining, available), outLimit - outPos);
            input.AsSpan(inPos, count).CopyTo(window.AsSpan(outPos));
            inPos += count;
            outPos += count;
            _storedRemaining -= count;
            return true;
        }

        private bool TryDecodeCoded(byte[] input, ref int inPos, int inEnd, bool inputEnded, byte[] window, ref int outPos, int outLimit)
        {
            // Past the real end the padding is zeros, and a valid stream never needs more than the
            // eight bytes a refill may have looked ahead.
            int inLimit = inputEnded ? inEnd + sizeof(ulong) : inEnd - LoopInputMargin;
            LoopStatus status = RunFastLoop(input, ref inPos, inLimit, window, ref outPos, outLimit);
            if (status == LoopStatus.BadOffset)
            {
                throw new InvalidDataException("A deflate match reaches before the start of the output.");
            }
            if (status == LoopStatus.BlockDone)
            {
                ThrowIfOverrun(inPos, inEnd, inputEnded);
                EndBlock();
                return true;
            }
            if (outPos >= outLimit)
            {
                ThrowIfOverrun(inPos, inEnd, inputEnded);
                return true;
            }
            ThrowIfTruncated(inputEnded);
            Unread(ref inPos);
            return false;
        }

        private void EndBlock()
        {
            _state = _finalBlock ? BlockState.Finished : BlockState.Header;
        }

        private void Unread(ref int inPos)
        {
            inPos -= _bitsleft >> 3;
            _bitsleft &= 7;
            _bitbuf &= (1UL << _bitsleft) - 1;
        }

        private void ThrowIfOverrun(int inPos, int inEnd, bool inputEnded)
        {
            if (inputEnded && inPos - (_bitsleft >> 3) > inEnd)
            {
                throw new InvalidDataException("The deflate stream is truncated.");
            }
        }

        private static void ThrowIfTruncated(bool inputEnded)
        {
            if (inputEnded)
            {
                throw new InvalidDataException("The deflate stream is truncated.");
            }
        }

        private void RefillBits(ReadOnlySpan<byte> input, ref int inPos)
        {
            _bitbuf |= BinaryPrimitives.ReadUInt64LittleEndian(input[inPos..]) << _bitsleft;
            inPos += (63 - _bitsleft) >> 3;
            _bitsleft |= 56;
        }

        private uint PopBits(int count)
        {
            uint value = (uint)(_bitbuf & ((1UL << count) - 1));
            _bitbuf >>= count;
            _bitsleft -= count;
            return value;
        }

        private void LoadStaticCodes()
        {
            if (_staticCodesLoaded)
            {
                return;
            }
            Span<byte> lens = stackalloc byte[InflateFormat.MaxLitlenSymbols + InflateFormat.MaxOffsetSymbols];
            lens[..144].Fill(8);
            lens[144..256].Fill(9);
            lens[256..280].Fill(7);
            lens[280..288].Fill(8);
            lens[288..].Fill(5);
            BuildCodeTables(lens, InflateFormat.MaxLitlenSymbols, InflateFormat.MaxOffsetSymbols);
            _staticCodesLoaded = true;
        }

        private void ReadDynamicCodes(ReadOnlySpan<byte> input, ref int inPos)
        {
            _staticCodesLoaded = false;
            int litlenSymbols = (int)PopBits(5) + 257;
            int offsetSymbols = (int)PopBits(5) + 1;
            int precodeSymbols = (int)PopBits(4) + 4;

            Span<byte> precodeLens = stackalloc byte[InflateFormat.PrecodeSymbols];
            precodeLens.Clear();
            for (int i = 0; i < precodeSymbols; i++)
            {
                if (_bitsleft < 3)
                {
                    RefillBits(input, ref inPos);
                }
                precodeLens[InflateFormat.PrecodeOrder[i]] = (byte)PopBits(3);
            }

            Span<uint> precode = stackalloc uint[InflateFormat.PrecodeTableSize];
            Span<ushort> sorted = stackalloc ushort[InflateFormat.PrecodeSymbols];
            if (!HuffmanTableBuilder.TryBuild(
                precode, precodeLens, InflateFormat.PrecodeResults, InflateFormat.PrecodeTableBits,
                InflateFormat.MaxPrecodeLength, sorted, shrink: false, out _))
            {
                throw new InvalidDataException("A deflate block has an invalid code-length code.");
            }

            Span<byte> lens = stackalloc byte[LensSize];
            ReadCodeLengths(input, ref inPos, precode, lens, litlenSymbols + offsetSymbols);
            BuildCodeTables(lens[..(litlenSymbols + offsetSymbols)], litlenSymbols, offsetSymbols);
        }

        private void ReadCodeLengths(ReadOnlySpan<byte> input, ref int inPos, ReadOnlySpan<uint> precode, Span<byte> lens, int total)
        {
            int count = 0;
            while (count < total)
            {
                if (_bitsleft < InflateFormat.MaxPrecodeLength + 7)
                {
                    RefillBits(input, ref inPos);
                }
                uint entry = precode[(int)(_bitbuf & PrecodeMask)];
                _ = PopBits((int)(entry & 0xFF));
                uint symbol = entry >> 16;
                if (symbol < 16)
                {
                    lens[count++] = (byte)symbol;
                    continue;
                }
                int repeat = ReadRepeat(symbol, count, lens, out byte value);
                lens.Slice(count, repeat).Fill(value);
                count += repeat;
            }
            if (count != total)
            {
                throw new InvalidDataException("A deflate block's code lengths overrun its header.");
            }
        }

        private int ReadRepeat(uint symbol, int count, ReadOnlySpan<byte> lens, out byte value)
        {
            value = 0;
            if (symbol == 17)
            {
                return 3 + (int)PopBits(3);
            }
            if (symbol == 18)
            {
                return 11 + (int)PopBits(7);
            }
            if (count == 0)
            {
                throw new InvalidDataException("A deflate block repeats a code length before defining one.");
            }
            value = lens[count - 1];
            return 3 + (int)PopBits(2);
        }

        private void BuildCodeTables(ReadOnlySpan<byte> lens, int litlenSymbols, int offsetSymbols)
        {
            Span<ushort> sorted = stackalloc ushort[InflateFormat.MaxLitlenSymbols];
            if (!HuffmanTableBuilder.TryBuild(
                _offset.AsSpan(0, InflateFormat.OffsetTableSize), lens.Slice(litlenSymbols, offsetSymbols), InflateFormat.OffsetResults,
                InflateFormat.OffsetTableBits, InflateFormat.MaxCodewordLength, sorted, shrink: false, out _))
            {
                throw new InvalidDataException("A deflate block has an invalid offset code.");
            }
            if (!HuffmanTableBuilder.TryBuild(
                _litlen.AsSpan(0, InflateFormat.LitlenTableSize), lens[..litlenSymbols], InflateFormat.LitlenResults,
                InflateFormat.LitlenTableBits, InflateFormat.MaxCodewordLength, sorted, shrink: true, out int litlenBits))
            {
                throw new InvalidDataException("A deflate block has an invalid literal/length code.");
            }
            _litlenMask = (1UL << litlenBits) - 1;
        }
    }
}
