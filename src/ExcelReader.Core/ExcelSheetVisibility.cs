namespace ExcelReader.Core
{
    /// <summary>
    /// Whether a sheet is shown in the workbook's tab bar. The values match the on-disk encoding every
    /// format uses for it (XLSX's <c>state</c> attribute, XLSB's <c>BrtBundleSh.hsState</c>, XLS's
    /// <c>BoundSheet8.hsState</c>).
    /// </summary>
    public enum ExcelSheetVisibility
    {
        /// <summary>Shown in the tab bar. What a sheet is when the format says nothing about it.</summary>
        Visible = 0,
        /// <summary>Hidden, and unhideable through the application's own UI.</summary>
        Hidden = 1,
        /// <summary>Hidden, and not offered by the application's unhide dialog — only a macro or a library like this one reveals it.</summary>
        VeryHidden = 2
    }
}
