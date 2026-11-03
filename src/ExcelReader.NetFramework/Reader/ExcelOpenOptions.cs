using ExcelReader.NetFramework.Parser.Internal;

namespace ExcelReader.NetFramework.Reader
{
    /// <summary>Options for opening a workbook. A null value keeps the native default.</summary>
    public sealed class ExcelOpenOptions
    {
        /// <summary>The file format. <see cref="ExcelFormat.Auto"/> detects it.</summary>
        public ExcelFormat Format { get; set; }

        /// <summary>The password of an encrypted workbook.</summary>
        public string? Password { get; set; }

        /// <summary>The CSV field delimiter. Must be an ASCII character.</summary>
        public char? CsvDelimiter { get; set; }

        /// <summary>The CSV quote character. Must be an ASCII character.</summary>
        public char? CsvQuote { get; set; }

        /// <summary>Whether a leading byte order mark is detected in CSV input.</summary>
        public bool? CsvDetectBom { get; set; }
    }
}
