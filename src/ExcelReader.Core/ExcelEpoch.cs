namespace ExcelReader.Core
{
    // Excel reserves serial 60 for the fictitious 1900-02-29; OADate does not. Keep this conversion
    // shared so the binary writers and the public Cell date reader cannot drift at the boundary.
    internal static class ExcelEpoch
    {
        internal static double SerialToOADate(double serial, bool date1904)
        {
            if (date1904)
            {
                return serial + 1462.0;
            }
            var isLessThan60 = serial < 60.0;
            return isLessThan60 ? serial + 1.0 : serial;
        }

        private static void ThrowIfNegative(double value, string paramName, string message)
        {
            if (value >= 0.0)
            {
                return;
            }
            throw new ArgumentOutOfRangeException(paramName, value, message);
        }

        internal static double OADateToSerial(double oadate, bool date1904)
        {
            double serial = oadate;
            if (date1904)
            {
                serial -= 1462.0;
                ThrowIfNegative(serial, nameof(oadate), "Dates before the workbook epoch cannot be written to Excel.");
                return serial;
            }
            if (oadate < 61.0)
            {
                serial--;
            }
            ThrowIfNegative(serial, nameof(oadate), "Dates before the workbook epoch cannot be written to Excel.");
            return serial;
        }
    }
}
