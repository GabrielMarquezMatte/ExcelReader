namespace ExcelReader.Core.ValueObjects
{
    internal static class FastDate
    {
        private const int MaxFractionDigits = 7;

        private static readonly long[] FractionScale = [1_000_000, 100_000, 10_000, 1_000, 100, 10, 1];

        public static bool TryParse(ReadOnlySpan<byte> s, out DateTime value)
        {
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
