namespace ExcelReader.Core.Writer
{
    /// <summary>Configures an Excel table written with <see cref="ITableSheetWriter{TRow}.BeginTable"/>.</summary>
    public sealed record ExcelTableOptions
    {
        /// <summary>The zero-based column of the table's first column. Defaults to 0 (column A).</summary>
        /// <remarks>Rows written while the table is open must skip to this column themselves.</remarks>
        public int FirstColumn { get; init; }

        /// <summary>
        /// The built-in table style: <c>TableStyleLight1</c>–<c>21</c>, <c>TableStyleMedium1</c>–<c>28</c> or
        /// <c>TableStyleDark1</c>–<c>11</c>; <see langword="null"/> for no style. Defaults to <c>TableStyleMedium2</c>.
        /// </summary>
        public string? StyleName { get; init; } = "TableStyleMedium2";

        /// <summary>Whether alternate rows are banded. Defaults to <see langword="true"/>.</summary>
        public bool ShowRowStripes { get; init; } = true;

        /// <summary>Whether alternate columns are banded. Defaults to <see langword="false"/>.</summary>
        public bool ShowColumnStripes { get; init; }

        /// <summary>Whether the first column gets the style's emphasis. Defaults to <see langword="false"/>.</summary>
        public bool ShowFirstColumn { get; init; }

        /// <summary>Whether the last column gets the style's emphasis. Defaults to <see langword="false"/>.</summary>
        public bool ShowLastColumn { get; init; }

        /// <summary>The default options: column A, <c>TableStyleMedium2</c>, banded rows.</summary>
        public static ExcelTableOptions Default { get; } = new();
    }
}
