namespace ExcelReader.Core.Writer.Internal
{
    // Shared BIFF8 (.xls) / BIFF12 (.xlsb) date-serial epoch handling: both formats store dates as
    // OADate-style serials, offset by 1462 days for workbooks using the 1904 date system.
    internal static class DateSerial
    {
        internal static double ForEpoch(double serial, bool date1904)
        {
            return date1904 ? serial - 1462.0 : serial;
        }
    }
}
