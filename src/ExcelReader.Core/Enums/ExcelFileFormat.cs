namespace ExcelReader.Core.Enums
{
    /// <summary>
    /// Identifies the on-disk Excel file format a workbook was read from or is to be written as.
    /// </summary>
    public enum ExcelFileFormat
    {
        /// <summary>The format could not be determined.</summary>
        Unknown,
        /// <summary>The legacy binary Excel format (.xls).</summary>
        Xls,
        /// <summary>The Office Open XML format (.xlsx).</summary>
        Xlsx,
        /// <summary>The Excel binary workbook format (.xlsb).</summary>
        Xlsb
    }
}
