using System.Runtime.InteropServices;
using System.Runtime.Intrinsics;

namespace ExcelReader.Core.ValueObjects
{
    internal static class FastDate
    {
        private const int MaxFractionDigits = 7;
        private const int RoundTripLength = 27;

        private static readonly long[] FractionScale = [1_000_000, 100_000, 10_000, 1_000, 100, 10, 1];

        private const uint HeadDigitPositions = 0xD96F;
        private const uint TailDigitPositions = 0xFEDB;

        private static readonly Vector128<byte> DateGather = Vector128.Create(
            0, 1, 2, 3, 5, 6, 8, 9, 0x80, 0x80, 0x80, 0x80, 0x80, 0x80, 0x80, 0x80);

        private static readonly Vector128<byte> TimeGather = Vector128.Create(
            0, 1, 3, 4, 6, 7, 0x80, 0x80, 9, 10, 11, 12, 13, 14, 15, 0x80);

        public static bool TryParse(ReadOnlySpan<byte> s, out DateTime value)
        {
            if (s.Length == RoundTripLength && Vector128.IsHardwareAccelerated)
            {
                return TryParseRoundTrip(s, out value);
            }
            value = default;
            if (s.Length < 10 || s[4] != (byte)'-' || s[7] != (byte)'-')
            {
                return false;
            }
            if (!TryDigits(s, 0, 4, out int year)
                || !TryTwoDigits(s, 5, out int month)
                || !TryTwoDigits(s, 8, out int day))
            {
                return false;
            }
            if (year < 1 || month is < 1 or > 12 || day < 1 || day > DateTime.DaysInMonth(year, month))
            {
                return false;
            }

            var date = new DateTime(year, month, day);
            if (s.Length == 10)
            {
                value = date;
                return true;
            }

            if (s.Length < 19 || s[10] is not ((byte)'T' or (byte)' ') || s[13] != (byte)':' || s[16] != (byte)':')
            {
                return false;
            }
            if (!TryTwoDigits(s, 11, out int hour)
                || !TryTwoDigits(s, 14, out int minute)
                || !TryTwoDigits(s, 17, out int second))
            {
                return false;
            }
            if (hour > 23 || minute > 59 || second > 59)
            {
                return false;
            }

            long ticks = date.Ticks
                + (hour * TimeSpan.TicksPerHour)
                + (minute * TimeSpan.TicksPerMinute)
                + (second * TimeSpan.TicksPerSecond);

            if (s.Length > 19)
            {
                if (s[19] != (byte)'.' || !TryFraction(s[20..], out long fraction))
                {
                    return false;
                }
                ticks += fraction;
            }

            value = new DateTime(ticks);
            return true;
        }

        private static bool TryParseRoundTrip(ReadOnlySpan<byte> s, out DateTime value)
        {
            value = default;
            if (s[4] != (byte)'-' || s[7] != (byte)'-' || s[13] != (byte)':'
                || s[16] != (byte)':' || s[19] != (byte)'.'
                || s[10] is not ((byte)'T' or (byte)' '))
            {
                return false;
            }

            ref byte origin = ref MemoryMarshal.GetReference(s);
            Vector128<byte> nine = Vector128.Create((byte)9);
            Vector128<byte> ascii = Vector128.Create((byte)'0');
            Vector128<byte> head = Vector128.LoadUnsafe(ref origin, 0) - ascii;
            Vector128<byte> tail = Vector128.LoadUnsafe(ref origin, 11) - ascii;

            if ((Vector128.LessThanOrEqual(head, nine).ExtractMostSignificantBits() & HeadDigitPositions) != HeadDigitPositions
                || (Vector128.LessThanOrEqual(tail, nine).ExtractMostSignificantBits() & TailDigitPositions) != TailDigitPositions)
            {
                return false;
            }

            ulong date = Vector128.Shuffle(head, DateGather).AsUInt64().GetElement(0);
            Vector128<ulong> rest = Vector128.Shuffle(tail, TimeGather).AsUInt64();

            ulong datePairs = FoldPairs(date);
            int year = (int)(((datePairs & 0xFF) * 100) + ((datePairs >> 16) & 0xFF));
            int month = (int)((datePairs >> 32) & 0xFF);
            int day = (int)((datePairs >> 48) & 0xFF);

            ulong clockPairs = FoldPairs(rest.GetElement(0));
            int hour = (int)(clockPairs & 0xFF);
            int minute = (int)((clockPairs >> 16) & 0xFF);
            int second = (int)((clockPairs >> 32) & 0xFF);

            if (year < 1 || month is < 1 or > 12 || day < 1 || day > DateTime.DaysInMonth(year, month)
                || hour > 23 || minute > 59 || second > 59)
            {
                return false;
            }

            ulong fractionPairs = FoldPairs(rest.GetElement(1));
            ulong fractionQuads = ((fractionPairs * 100) + (fractionPairs >> 16)) & 0x0000FFFF0000FFFFUL;
            ulong fraction = (((fractionQuads * 10000) + (fractionQuads >> 32)) & 0xFFFFFFFFUL) / 10;

            value = new DateTime(
                new DateTime(year, month, day).Ticks
                + (hour * TimeSpan.TicksPerHour)
                + (minute * TimeSpan.TicksPerMinute)
                + (second * TimeSpan.TicksPerSecond)
                + (long)fraction);
            return true;
        }

        private static ulong FoldPairs(ulong digits)
        {
            return ((digits * 10) + (digits >> 8)) & 0x00FF00FF00FF00FFUL;
        }

        private static bool TryFraction(ReadOnlySpan<byte> s, out long ticks)
        {
            ticks = 0;
            if (s.Length is 0 or > MaxFractionDigits)
            {
                return false;
            }
            for (int i = 0; i < s.Length; i++)
            {
                uint d = (uint)(s[i] - '0');
                if (d > 9)
                {
                    return false;
                }
                ticks += d * FractionScale[i];
            }
            return true;
        }

        private static bool TryTwoDigits(ReadOnlySpan<byte> s, int offset, out int value)
        {
            uint hi = (uint)(s[offset] - '0');
            uint lo = (uint)(s[offset + 1] - '0');
            value = (int)((hi * 10) + lo);
            return hi <= 9 && lo <= 9;
        }

        private static bool TryDigits(ReadOnlySpan<byte> s, int offset, int count, out int value)
        {
            int v = 0;
            for (int i = offset; i < offset + count; i++)
            {
                uint d = (uint)(s[i] - '0');
                if (d > 9)
                {
                    value = 0;
                    return false;
                }
                v = (v * 10) + (int)d;
            }
            value = v;
            return true;
        }
    }
}
