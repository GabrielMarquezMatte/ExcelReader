namespace ExcelReader.Core.Enums
{
    /// <summary>
    /// The byte sequence a <see cref="Writer.CsvWriter"/> writes after each record.
    /// </summary>
    public enum CsvNewLine
    {
        /// <summary>CR LF (<c>\r\n</c>), what RFC 4180 specifies and Excel writes.</summary>
        CarriageReturnLineFeed = 0,
        /// <summary>LF (<c>\n</c>), for Unix-oriented consumers.</summary>
        LineFeed = 1
    }
}
