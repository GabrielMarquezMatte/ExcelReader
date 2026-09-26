using System.Numerics;
using System.Runtime.CompilerServices;

namespace ExcelReader.Core.Reader.Internal
{
    internal static class FastDouble
    {
        private const int MaxExponentDigits = 3;
        private const int MaxMantissaDigits = 19;
        private const int MaxLosslessDigits = 15;
        private const int MinPowerOfFive = -64;
        private const int MaxPowerOfFive = 64;

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        public static bool TryParse(ReadOnlySpan<byte> s, out double value) => TryParse(s, MaxMantissaDigits, out value);

        // Rejects text with more significant digits than a double round-trips, so a caller that keeps
        // the double in place of the text never loses digits a decimal or long parse of the text would see.
        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        public static bool TryParseLossless(ReadOnlySpan<byte> s, out double value) => TryParse(s, MaxLosslessDigits, out value);

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        private static bool TryParse(ReadOnlySpan<byte> s, int maxDigits, out double value)
        {
            value = 0;
            if (s.IsEmpty)
            {
                return false;
            }
            int i = 0;
            bool neg = s[0] == (byte)'-';
            if (neg || s[0] == (byte)'+')
            {
                i = 1;
            }

            // Integer digits, then fraction digits, each in its own tight loop: one loop over both paid a
            // mispredicted branch at the dot and on every digit's dot/leading-zero checks.
            ulong mantissa = 0;
            int integerStart = i;
            for (; i < s.Length; i++)
            {
                uint d = (uint)(s[i] - (byte)'0');
                if (d > 9)
                {
                    break;
                }
                mantissa = (mantissa * 10) + d;
            }
            int integerDigits = i - integerStart;
            int scale = 0;
            if (i < s.Length && s[i] == (byte)'.')
            {
                int fractionStart = ++i;
                for (; i < s.Length; i++)
                {
                    uint d = (uint)(s[i] - (byte)'0');
                    if (d > 9)
                    {
                        break;
                    }
                    mantissa = (mantissa * 10) + d;
                }
                scale = i - fractionStart;
            }
            int allDigits = integerDigits + scale;
            if (allDigits == 0 || (allDigits > maxDigits && SignificantDigits(s[integerStart..i]) > maxDigits))
            {
                return false;
            }
            int exponent = 0;
            if (i < s.Length && (s[i] is not ((byte)'e' or (byte)'E') || !TryExponent(s[(i + 1)..], out exponent)))
            {
                return false;
            }

            scale -= exponent;
            if (scale is < -22 or > 22 || mantissa > 1UL << 53)
            {
                return EiselLemire(mantissa, -scale, neg, out value);
            }

            double result = mantissa;
            if (scale > 0)
            {
                result /= Pow10(scale);
            }
            else if (scale < 0)
            {
                result *= Pow10(-scale);
            }
            value = neg ? -result : result;
            return true;
        }

        // Eisel-Lemire as ported from fast_float's compute_float for binary64. Returns false for
        // subnormals, overflow and powers outside the table so the caller's double.TryParse decides.
        [MethodImpl(MethodImplOptions.NoInlining)]
        private static bool EiselLemire(ulong w, int q, bool neg, out double value)
        {
            value = 0;
            if (q is < MinPowerOfFive or > MaxPowerOfFive)
            {
                return false;
            }
            if (w == 0)
            {
                value = neg ? -0.0 : 0.0;
                return true;
            }
            int lz = BitOperations.LeadingZeroCount(w);
            w <<= lz;
            int index = 2 * (q - MinPowerOfFive);
            ReadOnlySpan<ulong> powers = PowersOfFive;
            ulong hi = Math.BigMul(w, powers[index], out ulong lo);
            const ulong PrecisionMask = ulong.MaxValue >> 55;
            if ((hi & PrecisionMask) == PrecisionMask)
            {
                ulong hi2 = Math.BigMul(w, powers[index + 1], out _);
                lo += hi2;
                if (hi2 > lo)
                {
                    hi++;
                }
                if (lo == ulong.MaxValue && q is < -27 or > 55)
                {
                    return false;
                }
            }
            int upperBit = (int)(hi >> 63);
            int shift = upperBit + 9;
            ulong m = hi >> shift;
            int biasedExponent = ((217706 * q) >> 16) + 63 + upperBit - lz + 1023;
            if (biasedExponent <= 0)
            {
                return false;
            }
            if (lo <= 1 && q is >= -4 and <= 23 && (m & 3) == 1 && (m << shift) == hi)
            {
                m &= ~1UL;
            }
            m += m & 1;
            m >>= 1;
            if (m >= 2UL << 52)
            {
                m = 1UL << 52;
                biasedExponent++;
            }
            m &= ~(1UL << 52);
            if (biasedExponent >= 0x7FF)
            {
                return false;
            }
            value = BitConverter.UInt64BitsToDouble(m | ((ulong)biasedExponent << 52) | (neg ? 1UL << 63 : 0));
            return true;
        }

        // Leading zeros add nothing to the mantissa, so they neither overflow it nor count toward the limit.
        [MethodImpl(MethodImplOptions.NoInlining)]
        private static int SignificantDigits(ReadOnlySpan<byte> number)
        {
            int digits = 0;
            foreach (byte c in number)
            {
                if (c != (byte)'.' && (digits > 0 || c != (byte)'0'))
                {
                    digits++;
                }
            }
            return digits;
        }

        [MethodImpl(MethodImplOptions.NoInlining)]
        private static bool TryExponent(ReadOnlySpan<byte> s, out int exponent)
        {
            exponent = 0;
            if (s.IsEmpty)
            {
                return false;
            }
            int i = 0;
            bool neg = s[0] == (byte)'-';
            if (neg || s[0] == (byte)'+')
            {
                i = 1;
            }
            if (i >= s.Length || s.Length - i > MaxExponentDigits)
            {
                return false;
            }
            int value = 0;
            for (; i < s.Length; i++)
            {
                uint d = (uint)(s[i] - (byte)'0');
                if (d > 9)
                {
                    return false;
                }
                value = (value * 10) + (int)d;
            }
            exponent = neg ? -value : value;
            return true;
        }

        private static double Pow10(int scale)
        {
            return PowersOfTen[scale];
        }

        private static ReadOnlySpan<double> PowersOfTen =>
        [
            1e0, 1e1, 1e2, 1e3, 1e4, 1e5, 1e6, 1e7, 1e8, 1e9, 1e10, 1e11,
            1e12, 1e13, 1e14, 1e15, 1e16, 1e17, 1e18, 1e19, 1e20, 1e21, 1e22,
        ];

        // ponytail: fast_float's 128-bit truncated 5^q table cut to q in [-64, 64]; widen to its full
        // [-342, 308] if inputs outside that range ever show up often enough to matter.
        private static ReadOnlySpan<ulong> PowersOfFive =>
        [
            0xA87FEA27A539E9A5, 0x3F2398D747B36224, 0xD29FE4B18E88640E, 0x8EEC7F0D19A03AAD,
            0x83A3EEEEF9153E89, 0x1953CF68300424AC, 0xA48CEAAAB75A8E2B, 0x5FA8C3423C052DD7,
            0xCDB02555653131B6, 0x3792F412CB06794D, 0x808E17555F3EBF11, 0xE2BBD88BBEE40BD0,
            0xA0B19D2AB70E6ED6, 0x5B6ACEAEAE9D0EC4, 0xC8DE047564D20A8B, 0xF245825A5A445275,
            0xFB158592BE068D2E, 0xEED6E2F0F0D56712, 0x9CED737BB6C4183D, 0x55464DD69685606B,
            0xC428D05AA4751E4C, 0xAA97E14C3C26B886, 0xF53304714D9265DF, 0xD53DD99F4B3066A8,
            0x993FE2C6D07B7FAB, 0xE546A8038EFE4029, 0xBF8FDB78849A5F96, 0xDE98520472BDD033,
            0xEF73D256A5C0F77C, 0x963E66858F6D4440, 0x95A8637627989AAD, 0xDDE7001379A44AA8,
            0xBB127C53B17EC159, 0x5560C018580D5D52, 0xE9D71B689DDE71AF, 0xAAB8F01E6E10B4A6,
            0x9226712162AB070D, 0xCAB3961304CA70E8, 0xB6B00D69BB55C8D1, 0x3D607B97C5FD0D22,
            0xE45C10C42A2B3B05, 0x8CB89A7DB77C506A, 0x8EB98A7A9A5B04E3, 0x77F3608E92ADB242,
            0xB267ED1940F1C61C, 0x55F038B237591ED3, 0xDF01E85F912E37A3, 0x6B6C46DEC52F6688,
            0x8B61313BBABCE2C6, 0x2323AC4B3B3DA015, 0xAE397D8AA96C1B77, 0xABEC975E0A0D081A,
            0xD9C7DCED53C72255, 0x96E7BD358C904A21, 0x881CEA14545C7575, 0x7E50D64177DA2E54,
            0xAA242499697392D2, 0xDDE50BD1D5D0B9E9, 0xD4AD2DBFC3D07787, 0x955E4EC64B44E864,
            0x84EC3C97DA624AB4, 0xBD5AF13BEF0B113E, 0xA6274BBDD0FADD61, 0xECB1AD8AEACDD58E,
            0xCFB11EAD453994BA, 0x67DE18EDA5814AF2, 0x81CEB32C4B43FCF4, 0x80EACF948770CED7,
            0xA2425FF75E14FC31, 0xA1258379A94D028D, 0xCAD2F7F5359A3B3E, 0x096EE45813A04330,
            0xFD87B5F28300CA0D, 0x8BCA9D6E188853FC, 0x9E74D1B791E07E48, 0x775EA264CF55347E,
            0xC612062576589DDA, 0x95364AFE032A819E, 0xF79687AED3EEC551, 0x3A83DDBD83F52205,
            0x9ABE14CD44753B52, 0xC4926A9672793543, 0xC16D9A0095928A27, 0x75B7053C0F178294,
            0xF1C90080BAF72CB1, 0x5324C68B12DD6339, 0x971DA05074DA7BEE, 0xD3F6FC16EBCA5E04,
            0xBCE5086492111AEA, 0x88F4BB1CA6BCF585, 0xEC1E4A7DB69561A5, 0x2B31E9E3D06C32E6,
            0x9392EE8E921D5D07, 0x3AFF322E62439FD0, 0xB877AA3236A4B449, 0x09BEFEB9FAD487C3,
            0xE69594BEC44DE15B, 0x4C2EBE687989A9B4, 0x901D7CF73AB0ACD9, 0x0F9D37014BF60A11,
            0xB424DC35095CD80F, 0x538484C19EF38C95, 0xE12E13424BB40E13, 0x2865A5F206B06FBA,
            0x8CBCCC096F5088CB, 0xF93F87B7442E45D4, 0xAFEBFF0BCB24AAFE, 0xF78F69A51539D749,
            0xDBE6FECEBDEDD5BE, 0xB573440E5A884D1C, 0x89705F4136B4A597, 0x31680A88F8953031,
            0xABCC77118461CEFC, 0xFDC20D2B36BA7C3E, 0xD6BF94D5E57A42BC, 0x3D32907604691B4D,
            0x8637BD05AF6C69B5, 0xA63F9A49C2C1B110, 0xA7C5AC471B478423, 0x0FCF80DC33721D54,
            0xD1B71758E219652B, 0xD3C36113404EA4A9, 0x83126E978D4FDF3B, 0x645A1CAC083126EA,
            0xA3D70A3D70A3D70A, 0x3D70A3D70A3D70A4, 0xCCCCCCCCCCCCCCCC, 0xCCCCCCCCCCCCCCCD,
            0x8000000000000000, 0x0000000000000000, 0xA000000000000000, 0x0000000000000000,
            0xC800000000000000, 0x0000000000000000, 0xFA00000000000000, 0x0000000000000000,
            0x9C40000000000000, 0x0000000000000000, 0xC350000000000000, 0x0000000000000000,
            0xF424000000000000, 0x0000000000000000, 0x9896800000000000, 0x0000000000000000,
            0xBEBC200000000000, 0x0000000000000000, 0xEE6B280000000000, 0x0000000000000000,
            0x9502F90000000000, 0x0000000000000000, 0xBA43B74000000000, 0x0000000000000000,
            0xE8D4A51000000000, 0x0000000000000000, 0x9184E72A00000000, 0x0000000000000000,
            0xB5E620F480000000, 0x0000000000000000, 0xE35FA931A0000000, 0x0000000000000000,
            0x8E1BC9BF04000000, 0x0000000000000000, 0xB1A2BC2EC5000000, 0x0000000000000000,
            0xDE0B6B3A76400000, 0x0000000000000000, 0x8AC7230489E80000, 0x0000000000000000,
            0xAD78EBC5AC620000, 0x0000000000000000, 0xD8D726B7177A8000, 0x0000000000000000,
            0x878678326EAC9000, 0x0000000000000000, 0xA968163F0A57B400, 0x0000000000000000,
            0xD3C21BCECCEDA100, 0x0000000000000000, 0x84595161401484A0, 0x0000000000000000,
            0xA56FA5B99019A5C8, 0x0000000000000000, 0xCECB8F27F4200F3A, 0x0000000000000000,
            0x813F3978F8940984, 0x4000000000000000, 0xA18F07D736B90BE5, 0x5000000000000000,
            0xC9F2C9CD04674EDE, 0xA400000000000000, 0xFC6F7C4045812296, 0x4D00000000000000,
            0x9DC5ADA82B70B59D, 0xF020000000000000, 0xC5371912364CE305, 0x6C28000000000000,
            0xF684DF56C3E01BC6, 0xC732000000000000, 0x9A130B963A6C115C, 0x3C7F400000000000,
            0xC097CE7BC90715B3, 0x4B9F100000000000, 0xF0BDC21ABB48DB20, 0x1E86D40000000000,
            0x96769950B50D88F4, 0x1314448000000000, 0xBC143FA4E250EB31, 0x17D955A000000000,
            0xEB194F8E1AE525FD, 0x5DCFAB0800000000, 0x92EFD1B8D0CF37BE, 0x5AA1CAE500000000,
            0xB7ABC627050305AD, 0xF14A3D9E40000000, 0xE596B7B0C643C719, 0x6D9CCD05D0000000,
            0x8F7E32CE7BEA5C6F, 0xE4820023A2000000, 0xB35DBF821AE4F38B, 0xDDA2802C8A800000,
            0xE0352F62A19E306E, 0xD50B2037AD200000, 0x8C213D9DA502DE45, 0x4526F422CC340000,
            0xAF298D050E4395D6, 0x9670B12B7F410000, 0xDAF3F04651D47B4C, 0x3C0CDD765F114000,
            0x88D8762BF324CD0F, 0xA5880A69FB6AC800, 0xAB0E93B6EFEE0053, 0x8EEA0D047A457A00,
            0xD5D238A4ABE98068, 0x72A4904598D6D880, 0x85A36366EB71F041, 0x47A6DA2B7F864750,
            0xA70C3C40A64E6C51, 0x999090B65F67D924, 0xD0CF4B50CFE20765, 0xFFF4B4E3F741CF6D,
            0x82818F1281ED449F, 0xBFF8F10E7A8921A4, 0xA321F2D7226895C7, 0xAFF72D52192B6A0D,
            0xCBEA6F8CEB02BB39, 0x9BF4F8A69F764490, 0xFEE50B7025C36A08, 0x02F236D04753D5B4,
            0x9F4F2726179A2245, 0x01D762422C946590, 0xC722F0EF9D80AAD6, 0x424D3AD2B7B97EF5,
            0xF8EBAD2B84E0D58B, 0xD2E0898765A7DEB2, 0x9B934C3B330C8577, 0x63CC55F49F88EB2F,
            0xC2781F49FFCFA6D5, 0x3CBF6B71C76B25FB
        ];
    }
}
