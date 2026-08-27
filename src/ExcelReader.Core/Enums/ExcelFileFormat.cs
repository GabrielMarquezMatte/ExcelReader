namespace ExcelReader.Core.Enums
{
    /// <summary>
    /// Identifies the on-disk Excel file format a workbook was read from or is to be written as.
    /// </summary>
    public enum ExcelFileFormat
    {
        /// <summary>The format could not be determined.</summary>
        Unknown = 0,
        /// <summary>The legacy binary Excel format (.xls).</summary>
        Xls = 1,
        /// <summary>The Office Open XML format (.xlsx).</summary>
        Xlsx = 2,
        /// <summary>The Excel binary workbook format (.xlsb).</summary>
        Xlsb = 3,
        /// <summary>An encrypted Office Open XML workbook. The concrete format (XLSX or XLSB) is
        /// unknowable until it has been decrypted, so detection reports this instead.</summary>
        EncryptedOoxml = 4
    }
}
