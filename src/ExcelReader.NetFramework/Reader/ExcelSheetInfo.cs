namespace ExcelReader.NetFramework.Reader
{
    /// <summary>The name and position of a sheet.</summary>
    public sealed class ExcelSheetInfo
    {
        internal ExcelSheetInfo(int index, string name, ExcelSheetVisibility visibility)
        {
            Index = index;
            Name = name;
            Visibility = visibility;
        }

        /// <summary>The zero-based position of the sheet.</summary>
        public int Index { get; }

        /// <summary>The sheet name.</summary>
        public string Name { get; }

        /// <summary>Whether the sheet is shown.</summary>
        public ExcelSheetVisibility Visibility { get; }
    }
}
