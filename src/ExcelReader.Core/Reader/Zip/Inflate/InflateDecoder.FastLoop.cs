// Ported from libdeflate's decompressor (MIT, Eric Biggers); the notice is in THIRD-PARTY-NOTICES.md.
using System.Buffers.Binary;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.CompilerServices;
using System.Runtime.Intrinsics.X86;

namespace ExcelReader.Core.Reader.Zip.Inflate
{
    internal sealed partial class InflateDecoder
    {
        private enum LoopStatus
        {
            BlockDone,
            Limit,
            BadOffset,
        }

        private unsafe LoopStatus RunFastLoop(byte[] input, ref int inPos, int inLimit, byte[] window, ref int outPos, int outLimit)
        {
            fixed (byte* inBase = input)
            fixed (byte* outBase = window)
            fixed (uint* litlen = _litlen)
            fixed (uint* offsets = _offset)
            {
                byte* inNext = inBase + inPos;
                byte* outNext = outBase + outPos;
                LoopStatus status = FastLoop(ref inNext, inBase + inLimit, outBase, ref outNext, outBase + outLimit, litlen, offsets);
                inPos = (int)(inNext - inBase);
                outPos = (int)(outNext - outBase);
                return status;
            }
        }

        // Entry layout, shared with HuffmanTableBuilder: bits 0-7 hold the bits to consume (codeword
        // plus extra bits), bits 8-11 the codeword length alone, bits 16 and up the literal, the
        // base length or offset, or the subtable start.
        //
        // Bit budget: a refill leaves at least 56 bits. Three literals from the main table take at
        // most 33; a length takes at most 20 and an offset at most 28, hence the refill between them.
        [SuppressMessage("Design", "MA0051:Method is too long",
            Justification = "The decode loop is deliberately one call-free method: every live value has to stay in a register across it, and a call inside the loop spills them.")]
        private unsafe LoopStatus FastLoop(
            ref byte* inPosition, byte* inLimit, byte* window, ref byte* outPosition, byte* outLimit, uint* litlen, uint* offsets)
        {
            ulong bitbuf = _bitbuf;
            int bitsleft = _bitsleft;
            byte* inNext = inPosition;
            byte* outNext = outPosition;
            ulong mask = _litlenMask;
            LoopStatus status;

            bitbuf |= LoadWord(inNext) << bitsleft;
            inNext += (63 - bitsleft) >> 3;
            bitsleft |= 56;
            uint entry = litlen[bitbuf & mask];

            while (true)
            {
                if (inNext >= inLimit || outNext >= outLimit)
                {
                    status = LoopStatus.Limit;
                    break;
                }

                ulong saved = bitbuf;
                bitbuf >>= (int)entry;
                bitsleft -= (int)(entry & 0xFF);

                if ((int)entry < 0)
                {
                    uint literal = entry >> 16;
                    entry = litlen[bitbuf & mask];
                    saved = bitbuf;
                    bitbuf >>= (int)entry;
                    bitsleft -= (int)(entry & 0xFF);
                    *outNext++ = (byte)literal;
                    if ((int)entry < 0)
                    {
                        literal = entry >> 16;
                        entry = litlen[bitbuf & mask];
                        saved = bitbuf;
                        bitbuf >>= (int)entry;
                        bitsleft -= (int)(entry & 0xFF);
                        *outNext++ = (byte)literal;
                    }
                    if ((int)entry < 0)
                    {
                        *outNext++ = (byte)(entry >> 16);
                        entry = litlen[bitbuf & mask];
                        bitbuf |= LoadWord(inNext) << bitsleft;
                        inNext += (63 - bitsleft) >> 3;
                        bitsleft |= 56;
                        continue;
                    }
                }

                if ((entry & InflateFormat.Exceptional) != 0)
                {
                    if ((entry & InflateFormat.EndOfBlock) != 0)
                    {
                        status = LoopStatus.BlockDone;
                        break;
                    }
                    entry = litlen[(entry >> 16) + (uint)LowBits(bitbuf, (entry >> 8) & 0x3F)];
                    saved = bitbuf;
                    bitbuf >>= (int)entry;
                    bitsleft -= (int)(entry & 0xFF);
                    if ((int)entry < 0)
                    {
                        *outNext++ = (byte)(entry >> 16);
                        entry = litlen[bitbuf & mask];
                        bitbuf |= LoadWord(inNext) << bitsleft;
                        inNext += (63 - bitsleft) >> 3;
                        bitsleft |= 56;
                        continue;
                    }
                    if ((entry & InflateFormat.EndOfBlock) != 0)
                    {
                        status = LoopStatus.BlockDone;
                        break;
                    }
                }

                uint length = (entry >> 16) + (uint)(LowBits(saved, entry & 0xFF) >> (int)((entry >> 8) & 0xF));

                if (bitsleft < 30)
                {
                    bitbuf |= LoadWord(inNext) << bitsleft;
                    inNext += (63 - bitsleft) >> 3;
                    bitsleft |= 56;
                }
                uint offsetEntry = offsets[bitbuf & ((1 << InflateFormat.OffsetTableBits) - 1)];
                if ((offsetEntry & InflateFormat.Exceptional) != 0)
                {
                    bitbuf >>= InflateFormat.OffsetTableBits;
                    bitsleft -= InflateFormat.OffsetTableBits;
                    offsetEntry = offsets[(offsetEntry >> 16) + (uint)LowBits(bitbuf, (offsetEntry >> 8) & 0x3F)];
                }
                saved = bitbuf;
                bitbuf >>= (int)offsetEntry;
                bitsleft -= (int)(offsetEntry & 0xFF);
                uint offset = (offsetEntry >> 16) + (uint)(LowBits(saved, offsetEntry & 0xFF) >> (int)((offsetEntry >> 8) & 0xF));

                if (offset > (nuint)(outNext - window))
                {
                    status = LoopStatus.BadOffset;
                    break;
                }
                byte* source = outNext - offset;
                byte* destination = outNext;
                outNext += length;

                bitbuf |= LoadWord(inNext) << bitsleft;
                inNext += (63 - bitsleft) >> 3;
                bitsleft |= 56;
                entry = litlen[bitbuf & mask];

                CopyMatch(source, destination, outNext, offset);
            }

            _bitbuf = bitbuf;
            _bitsleft = bitsleft;
            inPosition = inNext;
            outPosition = outNext;
            return status;
        }

        // Copies forward a word at a time, so it may write up to 39 bytes past `end`; the caller's
        // output slack covers that. An offset below eight overlaps within one word and needs the
        // narrower strides.
        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        private static unsafe void CopyMatch(byte* source, byte* destination, byte* end, uint offset)
        {
            if (offset >= sizeof(ulong))
            {
                CopyWord(source, destination);
                CopyWord(source + 8, destination + 8);
                CopyWord(source + 16, destination + 16);
                CopyWord(source + 24, destination + 24);
                CopyWord(source + 32, destination + 32);
                source += 40;
                destination += 40;
                while (destination < end)
                {
                    CopyWord(source, destination);
                    CopyWord(source + 8, destination + 8);
                    source += 16;
                    destination += 16;
                }
                return;
            }
            if (offset == 1)
            {
                ulong repeated = 0x0101010101010101UL * *source;
                Unsafe.WriteUnaligned(destination, repeated);
                Unsafe.WriteUnaligned(destination + 8, repeated);
                destination += 16;
                while (destination < end)
                {
                    Unsafe.WriteUnaligned(destination, repeated);
                    destination += 8;
                }
                return;
            }
            do
            {
                CopyWord(source, destination);
                source += offset;
                destination += offset;
            }
            while (destination < end);
        }

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        private static unsafe void CopyWord(byte* source, byte* destination)
        {
            Unsafe.WriteUnaligned(destination, Unsafe.ReadUnaligned<ulong>(source));
        }

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        private static unsafe ulong LoadWord(byte* source)
        {
            ulong value = Unsafe.ReadUnaligned<ulong>(source);
            if (BitConverter.IsLittleEndian)
            {
                return value;
            }
            return BinaryPrimitives.ReverseEndianness(value);
        }

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        private static ulong LowBits(ulong value, uint count)
        {
            if (Bmi2.X64.IsSupported)
            {
                return Bmi2.X64.ZeroHighBits(value, count);
            }
            return value & ((1UL << (int)count) - 1);
        }
    }
}
