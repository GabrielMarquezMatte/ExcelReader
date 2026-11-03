namespace ExcelReader.NetFramework
{
    /// <summary>The file format of a workbook.</summary>
    public enum ExcelFormat
    {
        /// <summary>Detect the format from the content.</summary>
        Auto = 0,

        /// <summary>Legacy binary workbook.</summary>
        Xls = 1,

        /// <summary>Office Open XML workbook.</summary>
        Xlsx = 2,

        /// <summary>Binary Office Open XML workbook.</summary>
        Xlsb = 3,

        /// <summary>Delimited text. Never detected, it must be requested.</summary>
        Csv = 4,
    }
}
