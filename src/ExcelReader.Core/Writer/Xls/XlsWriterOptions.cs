namespace ExcelReader.Core.Writer.Xls
{
    /// <summary>Configures how a BIFF8 (.xls) workbook is written.</summary>
    public sealed record XlsWriterOptions
    {
        /// <summary>Whether the workbook uses the 1904 date system instead of the default 1900 system. Defaults to <see langword="false"/>.</summary>
        public bool Date1904 { get; init; }

        /// <summary>The default options: 1900 date system.</summary>
        public static XlsWriterOptions Default { get; } = new();
    }
}
