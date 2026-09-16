namespace ExcelReader.Core.ValueObjects
{
    internal static class FastDate
    {
        private const int MaxFractionDigits = 7;

        public static bool TryParse(ReadOnlySpan<byte> s, out DateTime value)
        {
            value = default;
            if (s.Length < 10 || s[4] != (byte)'-' || s[7] != (byte)'-')
            {
                return false;
            }
            if (!TryDigits(s, 0, 4, out int year)
                || !TryDigits(s, 5, 2, out int month)
                || !TryDigits(s, 8, 2, out int day))
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
            if (!TryDigits(s, 11, 2, out int hour)
                || !TryDigits(s, 14, 2, out int minute)
                || !TryDigits(s, 17, 2, out int second))
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

        // Rejects a trailing zone designator rather than guessing at it; those fall back to the
        // general parser, which is the only thing that knows what to do with an offset.
        private static bool TryFraction(ReadOnlySpan<byte> s, out long ticks)
        {
            ticks = 0;
            if (s.Length is 0 or > MaxFractionDigits)
            {
                return false;
            }
            long scale = TimeSpan.TicksPerSecond;
            for (int i = 0; i < s.Length; i++)
            {
                uint d = (uint)(s[i] - '0');
                if (d > 9)
                {
                    return false;
                }
                scale /= 10;
                ticks += d * scale;
            }
            return true;
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
