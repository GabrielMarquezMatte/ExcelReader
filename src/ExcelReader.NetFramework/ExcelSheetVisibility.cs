namespace ExcelReader.NetFramework
{
    /// <summary>Whether a sheet is shown in the workbook.</summary>
    public enum ExcelSheetVisibility
    {
        /// <summary>The sheet is shown.</summary>
        Visible = 0,

        /// <summary>The sheet is hidden and can be unhidden by the user.</summary>
        Hidden = 1,

        /// <summary>The sheet is hidden and cannot be unhidden from the user interface.</summary>
        VeryHidden = 2,
    }
}
