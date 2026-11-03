using ExcelReader.NetFramework.Interop;

namespace ExcelReader.NetFramework.Writer
{
    /// <summary>Writes a sheet as an Excel table (XLSX and XLSB only), its header row holding the column names.</summary>
    public sealed class ExcelTableOptions
    {
        /// <summary>The table name; required and unique in the workbook.</summary>
        public string Name { get; set; } = string.Empty;

        /// <summary>The table style name, or null for TableStyleMedium2. Cannot be combined with <see cref="NoStyle"/>.</summary>
        public string? Style { get; set; }

        /// <summary>Whether alternate rows are shaded. Defaults to true.</summary>
        public bool ShowRowStripes { get; set; } = true;

        /// <summary>Whether alternate columns are shaded.</summary>
        public bool ShowColumnStripes { get; set; }

        /// <summary>Whether the first column is emphasized.</summary>
        public bool EmphasizeFirstColumn { get; set; }

        /// <summary>Whether the last column is emphasized.</summary>
        public bool EmphasizeLastColumn { get; set; }

        /// <summary>Whether the table has no style at all.</summary>
        public bool NoStyle { get; set; }

        internal int Flags()
        {
            int flags = 0;
            if (NoStyle)
            {
                flags |= Xl.XL_TABLE_NO_STYLE;
            }

            if (!ShowRowStripes)
            {
                flags |= Xl.XL_TABLE_NO_ROW_STRIPES;
            }

            if (ShowColumnStripes)
            {
                flags |= Xl.XL_TABLE_COLUMN_STRIPES;
            }

            if (EmphasizeFirstColumn)
            {
                flags |= Xl.XL_TABLE_FIRST_COLUMN;
            }

            if (EmphasizeLastColumn)
            {
                flags |= Xl.XL_TABLE_LAST_COLUMN;
            }

            return flags;
        }
    }
}
