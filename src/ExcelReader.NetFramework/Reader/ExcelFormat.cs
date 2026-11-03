namespace ExcelReader.NetFramework.Reader
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

        /// <summary>
        /// Delimited text. <see cref="Auto"/> never picks it for a path or bytes, so request it there; for a stream,
        /// <see cref="Auto"/> reads anything without a ZIP or OLE2 signature as CSV.
        /// </summary>
        Csv = 4,
    }
}
