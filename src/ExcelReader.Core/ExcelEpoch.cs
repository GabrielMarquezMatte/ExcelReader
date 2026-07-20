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
            return serial switch
            {
                < 60.0 => serial + 1.0,
                60.0 => 60.0,
                _ => serial,
            };
        }

        internal static double OADateToSerial(double oadate, bool date1904)
        {
            double serial = oadate;
            if (date1904)
            {
                serial -= 1462.0;
            }
            else if (oadate < 61.0)
            {
                serial--;
            }
            if (serial < 0)
            {
                throw new ArgumentOutOfRangeException(nameof(oadate), "Dates before the workbook epoch cannot be written to Excel.");
            }
            return serial;
        }
    }
}
