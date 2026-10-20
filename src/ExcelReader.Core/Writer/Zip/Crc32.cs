using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.Intrinsics;
using System.Runtime.Intrinsics.X86;
using ArmCrc32 = System.Runtime.Intrinsics.Arm.Crc32;

namespace ExcelReader.Core.Writer.Zip
{
    /// <summary>
    /// The ZIP CRC-32: reflected polynomial 0xEDB88320, register inverted on entry and exit.
    /// <see cref="Append"/> chains, so <c>Append(Append(0, a), b)</c> equals <c>Append(0, a + b)</c>.
    /// </summary>
    /// <remarks>
    /// Long inputs are folded with carry-less multiplication, after Intel's "Fast CRC Computation for
    /// Generic Polynomials Using PCLMULQDQ Instruction". The folded 128-bit remainder is finished by
    /// running its 16 bytes through the byte table from a zero register.
    /// </remarks>
    internal static class Crc32
    {
        private const uint ReflectedPolynomial = 0xEDB8_8320;
        private const uint Polynomial = 0x04C1_1DB7;

        private static readonly uint[] Table = BuildTable();
        private static readonly uint[] SquaredPowersOfX = BuildSquaredPowersOfX();
        private static readonly Vector128<ulong> Fold128By128 = FoldConstants(128);
        private static readonly Vector128<ulong> Fold128By512 = FoldConstants(512);
        private static readonly Vector256<ulong> Fold256By256 = Vector256.Create(FoldConstants(256), FoldConstants(256));
        private static readonly Vector256<ulong> Fold256By1024 = Vector256.Create(FoldConstants(1024), FoldConstants(1024));

        internal static uint Append(uint crc, ReadOnlySpan<byte> data)
        {
            if (Pclmulqdq.V256.IsSupported && data.Length >= 256)
            {
                return AppendFold256(crc, data);
            }
            if (Pclmulqdq.IsSupported && data.Length >= 128)
            {
                return AppendFold128(crc, data);
            }
            if (ArmCrc32.Arm64.IsSupported)
            {
                return AppendArm64(crc, data);
            }
            return AppendTable(crc, data);
        }

        internal static uint AppendTable(uint crc, ReadOnlySpan<byte> data)
        {
            return ~TableUpdate(~crc, data);
        }

        internal static uint AppendArm64(uint crc, ReadOnlySpan<byte> data)
        {
            uint register = ~crc;
            ref byte source = ref MemoryMarshal.GetReference(data);
            int index = 0;
            for (; index + 8 <= data.Length; index += 8)
            {
                register = ArmCrc32.Arm64.ComputeCrc32(register, Unsafe.ReadUnaligned<ulong>(ref Unsafe.Add(ref source, index)));
            }
            for (; index < data.Length; index++)
            {
                register = ArmCrc32.ComputeCrc32(register, Unsafe.Add(ref source, index));
            }
            return ~register;
        }

        internal static uint AppendFold128(uint crc, ReadOnlySpan<byte> data)
        {
            if (data.Length < 64)
            {
                return AppendTable(crc, data);
            }
            uint register = Fold128(~crc, ref data);
            return ~TableUpdate(register, data);
        }

        internal static uint AppendFold256(uint crc, ReadOnlySpan<byte> data)
        {
            if (data.Length < 128)
            {
                return AppendTable(crc, data);
            }
            uint register = Fold256(~crc, ref data);
            return ~TableUpdate(register, data);
        }

        private static uint TableUpdate(uint register, ReadOnlySpan<byte> data)
        {
            ref uint table = ref MemoryMarshal.GetArrayDataReference(Table);
            foreach (byte value in data)
            {
                register = Unsafe.Add(ref table, (byte)(register ^ value)) ^ (register >> 8);
            }
            return register;
        }

        private static uint Fold256(uint register, ref ReadOnlySpan<byte> data)
        {
            ref byte source = ref MemoryMarshal.GetReference(data);
            Vector256<ulong> by1024 = Fold256By1024;
            Vector256<ulong> by256 = Fold256By256;
            Vector256<ulong> x0 = Load256(ref source, 0) ^ Vector256.CreateScalar((ulong)register);
            Vector256<ulong> x1 = Load256(ref source, 32);
            Vector256<ulong> x2 = Load256(ref source, 64);
            Vector256<ulong> x3 = Load256(ref source, 96);
            int position = 128;
            while (data.Length - position >= 128)
            {
                x0 = Fold(x0, by1024) ^ Load256(ref source, position);
                x1 = Fold(x1, by1024) ^ Load256(ref source, position + 32);
                x2 = Fold(x2, by1024) ^ Load256(ref source, position + 64);
                x3 = Fold(x3, by1024) ^ Load256(ref source, position + 96);
                position += 128;
            }
            x1 ^= Fold(x0, by256);
            x2 ^= Fold(x1, by256);
            x3 ^= Fold(x2, by256);
            Vector128<ulong> folded = Fold(x3.GetLower(), Fold128By128) ^ x3.GetUpper();
            return Finish(folded, ref source, position, ref data);
        }

        private static uint Fold128(uint register, ref ReadOnlySpan<byte> data)
        {
            ref byte source = ref MemoryMarshal.GetReference(data);
            Vector128<ulong> by512 = Fold128By512;
            Vector128<ulong> by128 = Fold128By128;
            Vector128<ulong> x0 = Load128(ref source, 0) ^ Vector128.CreateScalar((ulong)register);
            Vector128<ulong> x1 = Load128(ref source, 16);
            Vector128<ulong> x2 = Load128(ref source, 32);
            Vector128<ulong> x3 = Load128(ref source, 48);
            int position = 64;
            while (data.Length - position >= 64)
            {
                x0 = Fold(x0, by512) ^ Load128(ref source, position);
                x1 = Fold(x1, by512) ^ Load128(ref source, position + 16);
                x2 = Fold(x2, by512) ^ Load128(ref source, position + 32);
                x3 = Fold(x3, by512) ^ Load128(ref source, position + 48);
                position += 64;
            }
            x1 ^= Fold(x0, by128);
            x2 ^= Fold(x1, by128);
            Vector128<ulong> folded = x3 ^ Fold(x2, by128);
            return Finish(folded, ref source, position, ref data);
        }

        private static uint Finish(Vector128<ulong> folded, ref byte source, int position, ref ReadOnlySpan<byte> data)
        {
            while (data.Length - position >= 16)
            {
                folded = Fold(folded, Fold128By128) ^ Load128(ref source, position);
                position += 16;
            }
            data = data[position..];
            // Not a stackalloc: with one, this loop would be compiled fully optimized on first call instead of at tier 0.
            return TableUpdate(0, MemoryMarshal.AsBytes(new ReadOnlySpan<Vector128<ulong>>(in folded)));
        }

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        private static Vector256<ulong> Fold(Vector256<ulong> lanes, Vector256<ulong> constants)
        {
            return Pclmulqdq.V256.CarrylessMultiply(lanes, constants, 0x00) ^ Pclmulqdq.V256.CarrylessMultiply(lanes, constants, 0x11);
        }

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        private static Vector128<ulong> Fold(Vector128<ulong> lanes, Vector128<ulong> constants)
        {
            return Pclmulqdq.CarrylessMultiply(lanes, constants, 0x00) ^ Pclmulqdq.CarrylessMultiply(lanes, constants, 0x11);
        }

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        private static Vector256<ulong> Load256(ref byte source, int offset)
        {
            return Vector256.LoadUnsafe(ref Unsafe.Add(ref source, offset)).AsUInt64();
        }

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        private static Vector128<ulong> Load128(ref byte source, int offset)
        {
            return Vector128.LoadUnsafe(ref Unsafe.Add(ref source, offset)).AsUInt64();
        }

        // Moving a 128-bit lane forward by D bits multiplies its low qword by k(D + 32) and its high
        // qword by k(D - 32), where k(e) = reflect32(x^e mod P) << 1.
        private static Vector128<ulong> FoldConstants(int distanceBits)
        {
            return Vector128.Create(FoldConstant(distanceBits + 32), FoldConstant(distanceBits - 32));
        }

        private static ulong FoldConstant(int exponent)
        {
            uint remainder = 1;
            for (int i = 0; i < exponent; i++)
            {
                uint carry = 0u - (remainder >> 31);
                remainder = (remainder << 1) ^ (Polynomial & carry);
            }
            uint reflected = 0;
            for (int bit = 0; bit < 32; bit++)
            {
                reflected = (reflected << 1) | ((remainder >> bit) & 1);
            }
            return (ulong)reflected << 1;
        }

        private static uint[] BuildTable()
        {
            var table = new uint[256];
            for (uint index = 0; index < 256; index++)
            {
                uint entry = index;
                for (int bit = 0; bit < 8; bit++)
                {
                    entry = (entry >> 1) ^ (ReflectedPolynomial & (0u - (entry & 1)));
                }
                table[index] = entry;
            }
            return table;
        }

        /// <summary>
        /// The CRC of two byte runs laid end to end, from the CRC of each and the length of the second.
        /// </summary>
        internal static uint Combine(uint first, uint second, long secondLength)
        {
            ArgumentOutOfRangeException.ThrowIfNegative(secondLength);
            return MultiplyModP(PowerOfX(secondLength), first) ^ second;
        }

        // x^(8 * byteCount) mod P, by square and multiply over the bits of byteCount. Entry k of the table
        // is x^(2^k); the sequence repeats every 32 entries, and bytes start three squarings in.
        private static uint PowerOfX(long byteCount)
        {
            uint power = 1u << 31;
            int exponent = 3;
            while (byteCount != 0)
            {
                if ((byteCount & 1) != 0)
                {
                    power = MultiplyModP(SquaredPowersOfX[exponent & 31], power);
                }
                byteCount >>= 1;
                exponent++;
            }
            return power;
        }

        // Polynomial product mod P in the reflected domain, where bit 31 is x^0.
        private static uint MultiplyModP(uint a, uint b)
        {
            uint product = 0;
            for (uint mask = 1u << 31; mask != 0; mask >>= 1)
            {
                if ((a & mask) != 0)
                {
                    product ^= b;
                }
                b = (b >> 1) ^ (ReflectedPolynomial & (0u - (b & 1)));
            }
            return product;
        }

        private static uint[] BuildSquaredPowersOfX()
        {
            var table = new uint[32];
            uint power = 1u << 30;
            table[0] = power;
            for (int index = 1; index < table.Length; index++)
            {
                power = MultiplyModP(power, power);
                table[index] = power;
            }
            return table;
        }
    }
}
