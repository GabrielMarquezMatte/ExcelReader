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
        private static readonly Vector128<ulong> Fold128By128 = Vector128.Create(0x1_7519_97D0UL, 0x0_CCAA_009EUL);
        private static readonly Vector128<ulong> Fold128By512 = Vector128.Create(0x1_5444_2BD4UL, 0x1_C6E4_1596UL);
        private static readonly Vector256<ulong> Fold256By256 = Vector256.Create(0x0_F1DA_05AAUL, 0x1_5A54_6366UL, 0x0_F1DA_05AAUL, 0x1_5A54_6366UL);
        private static readonly Vector256<ulong> Fold256By1024 = Vector256.Create(0x1_E88E_F372UL, 0x1_4A7F_E880UL, 0x1_E88E_F372UL, 0x1_4A7F_E880UL);

        private static ReadOnlySpan<uint> Table =>
        [
            0x00000000, 0x77073096, 0xEE0E612C, 0x990951BA, 0x076DC419, 0x706AF48F, 0xE963A535, 0x9E6495A3,
            0x0EDB8832, 0x79DCB8A4, 0xE0D5E91E, 0x97D2D988, 0x09B64C2B, 0x7EB17CBD, 0xE7B82D07, 0x90BF1D91,
            0x1DB71064, 0x6AB020F2, 0xF3B97148, 0x84BE41DE, 0x1ADAD47D, 0x6DDDE4EB, 0xF4D4B551, 0x83D385C7,
            0x136C9856, 0x646BA8C0, 0xFD62F97A, 0x8A65C9EC, 0x14015C4F, 0x63066CD9, 0xFA0F3D63, 0x8D080DF5,
            0x3B6E20C8, 0x4C69105E, 0xD56041E4, 0xA2677172, 0x3C03E4D1, 0x4B04D447, 0xD20D85FD, 0xA50AB56B,
            0x35B5A8FA, 0x42B2986C, 0xDBBBC9D6, 0xACBCF940, 0x32D86CE3, 0x45DF5C75, 0xDCD60DCF, 0xABD13D59,
            0x26D930AC, 0x51DE003A, 0xC8D75180, 0xBFD06116, 0x21B4F4B5, 0x56B3C423, 0xCFBA9599, 0xB8BDA50F,
            0x2802B89E, 0x5F058808, 0xC60CD9B2, 0xB10BE924, 0x2F6F7C87, 0x58684C11, 0xC1611DAB, 0xB6662D3D,
            0x76DC4190, 0x01DB7106, 0x98D220BC, 0xEFD5102A, 0x71B18589, 0x06B6B51F, 0x9FBFE4A5, 0xE8B8D433,
            0x7807C9A2, 0x0F00F934, 0x9609A88E, 0xE10E9818, 0x7F6A0DBB, 0x086D3D2D, 0x91646C97, 0xE6635C01,
            0x6B6B51F4, 0x1C6C6162, 0x856530D8, 0xF262004E, 0x6C0695ED, 0x1B01A57B, 0x8208F4C1, 0xF50FC457,
            0x65B0D9C6, 0x12B7E950, 0x8BBEB8EA, 0xFCB9887C, 0x62DD1DDF, 0x15DA2D49, 0x8CD37CF3, 0xFBD44C65,
            0x4DB26158, 0x3AB551CE, 0xA3BC0074, 0xD4BB30E2, 0x4ADFA541, 0x3DD895D7, 0xA4D1C46D, 0xD3D6F4FB,
            0x4369E96A, 0x346ED9FC, 0xAD678846, 0xDA60B8D0, 0x44042D73, 0x33031DE5, 0xAA0A4C5F, 0xDD0D7CC9,
            0x5005713C, 0x270241AA, 0xBE0B1010, 0xC90C2086, 0x5768B525, 0x206F85B3, 0xB966D409, 0xCE61E49F,
            0x5EDEF90E, 0x29D9C998, 0xB0D09822, 0xC7D7A8B4, 0x59B33D17, 0x2EB40D81, 0xB7BD5C3B, 0xC0BA6CAD,
            0xEDB88320, 0x9ABFB3B6, 0x03B6E20C, 0x74B1D29A, 0xEAD54739, 0x9DD277AF, 0x04DB2615, 0x73DC1683,
            0xE3630B12, 0x94643B84, 0x0D6D6A3E, 0x7A6A5AA8, 0xE40ECF0B, 0x9309FF9D, 0x0A00AE27, 0x7D079EB1,
            0xF00F9344, 0x8708A3D2, 0x1E01F268, 0x6906C2FE, 0xF762575D, 0x806567CB, 0x196C3671, 0x6E6B06E7,
            0xFED41B76, 0x89D32BE0, 0x10DA7A5A, 0x67DD4ACC, 0xF9B9DF6F, 0x8EBEEFF9, 0x17B7BE43, 0x60B08ED5,
            0xD6D6A3E8, 0xA1D1937E, 0x38D8C2C4, 0x4FDFF252, 0xD1BB67F1, 0xA6BC5767, 0x3FB506DD, 0x48B2364B,
            0xD80D2BDA, 0xAF0A1B4C, 0x36034AF6, 0x41047A60, 0xDF60EFC3, 0xA867DF55, 0x316E8EEF, 0x4669BE79,
            0xCB61B38C, 0xBC66831A, 0x256FD2A0, 0x5268E236, 0xCC0C7795, 0xBB0B4703, 0x220216B9, 0x5505262F,
            0xC5BA3BBE, 0xB2BD0B28, 0x2BB45A92, 0x5CB36A04, 0xC2D7FFA7, 0xB5D0CF31, 0x2CD99E8B, 0x5BDEAE1D,
            0x9B64C2B0, 0xEC63F226, 0x756AA39C, 0x026D930A, 0x9C0906A9, 0xEB0E363F, 0x72076785, 0x05005713,
            0x95BF4A82, 0xE2B87A14, 0x7BB12BAE, 0x0CB61B38, 0x92D28E9B, 0xE5D5BE0D, 0x7CDCEFB7, 0x0BDBDF21,
            0x86D3D2D4, 0xF1D4E242, 0x68DDB3F8, 0x1FDA836E, 0x81BE16CD, 0xF6B9265B, 0x6FB077E1, 0x18B74777,
            0x88085AE6, 0xFF0F6A70, 0x66063BCA, 0x11010B5C, 0x8F659EFF, 0xF862AE69, 0x616BFFD3, 0x166CCF45,
            0xA00AE278, 0xD70DD2EE, 0x4E048354, 0x3903B3C2, 0xA7672661, 0xD06016F7, 0x4969474D, 0x3E6E77DB,
            0xAED16A4A, 0xD9D65ADC, 0x40DF0B66, 0x37D83BF0, 0xA9BCAE53, 0xDEBB9EC5, 0x47B2CF7F, 0x30B5FFE9,
            0xBDBDF21C, 0xCABAC28A, 0x53B39330, 0x24B4A3A6, 0xBAD03605, 0xCDD70693, 0x54DE5729, 0x23D967BF,
            0xB3667A2E, 0xC4614AB8, 0x5D681B02, 0x2A6F2B94, 0xB40BBE37, 0xC30C8EA1, 0x5A05DF1B, 0x2D02EF8D
        ];

        private static ReadOnlySpan<uint> SquaredPowersOfX =>
        [
            0x40000000, 0x20000000, 0x08000000, 0x00800000, 0x00008000, 0xEDB88320, 0xB1E6B092, 0xA06A2517,
            0xED627DAE, 0x88D14467, 0xD7BBFE6A, 0xEC447F11, 0x8E7EA170, 0x6427800E, 0x4D47BAE0, 0x09FE548F,
            0x83852D0F, 0x30362F1A, 0x7B5A9CC3, 0x31FEC169, 0x9FEC022A, 0x6C8DEDC4, 0x15D6874D, 0x5FDE7A4E,
            0xBAD90E37, 0x2E4E5EEF, 0x4EABA214, 0xA8A472C0, 0x429A969E, 0x148D302A, 0xC40BA6D0, 0xC4E22C3C
        ];

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
            ref uint table = ref MemoryMarshal.GetReference(Table);
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
    }
}
