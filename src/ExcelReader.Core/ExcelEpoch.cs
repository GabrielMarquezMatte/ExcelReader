namespace ExcelReader.Core
{
    // Excel reserves serial 60 for the fictitious 1900-02-29; OADate does not. Keep this conversion
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

        internal static double OADateToSerial(double oadate, bool date1904)
        {
            if (!TryOADateToSerial(oadate, date1904, out double serial))
            {
                throw new ArgumentOutOfRangeException(nameof(oadate), serial, "Dates before the workbook epoch cannot be written to Excel.");
            }
            return serial;
        }

        /// <summary>False for a date before the workbook epoch, which has no serial.</summary>
        internal static bool TryOADateToSerial(double oadate, bool date1904, out double serial)
        {
            if (date1904)
            {
                serial = oadate - 1462.0;
            }
            else
            {
                serial = oadate < 61.0 ? oadate - 1.0 : oadate;
            }
            return serial >= 0.0;
        }
    }
}
