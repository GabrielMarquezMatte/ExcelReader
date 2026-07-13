namespace ExcelReader.Core.Writer.Internal
{
    // Shared BIFF8 (.xls) / BIFF12 (.xlsb) date-serial epoch handling: both formats store dates as
    // OADate-style serials, offset by 1462 days for workbooks using the 1904 date system.
    internal static class DateSerial
    {
        internal static double ForEpoch(double serial, bool date1904)
        {
            // OADate is one day ahead of Excel for Jan 1–Feb 28 1900 because Excel
            // reserves serial 60 for its fictitious 1900-02-29.
            if (!date1904 && serial < 61.0)
            {
                serial -= 1.0;
            }
            if (date1904)
            {
                serial -= 1462.0;
            }
            if (serial < 0)
            {
                throw new ArgumentOutOfRangeException(nameof(serial), "Dates before the workbook epoch cannot be written to Excel.");
            }
            return serial;
        }
    }
}
