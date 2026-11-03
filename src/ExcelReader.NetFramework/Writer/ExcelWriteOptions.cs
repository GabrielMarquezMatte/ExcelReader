using ExcelReader.NetFramework.Reader;

namespace ExcelReader.NetFramework.Writer
{
    /// <summary>Options for writing a workbook. A null value keeps the native default.</summary>
    public sealed class ExcelWriteOptions
    {
        /// <summary>The file format. Null takes it from the path extension; <see cref="ExcelFormat.Auto"/> is not a writable format.</summary>
        public ExcelFormat? Format { get; set; }

        /// <summary>The CSV field delimiter. Must be an ASCII character.</summary>
        public char? CsvDelimiter { get; set; }

        /// <summary>The CSV quote character. Must be an ASCII character.</summary>
        public char? CsvQuote { get; set; }

        /// <summary>Whether dates use the 1904 epoch.</summary>
        public bool? Date1904 { get; set; }

        /// <summary>Whether strings go through the shared string table.</summary>
        public bool? UseSharedStrings { get; set; }

        /// <summary>The number of threads used to encode a sheet. Zero keeps the native default.</summary>
        public int DegreeOfParallelism { get; set; }

        /// <summary>The rows encoded per batch. Zero keeps the native default.</summary>
        public int BatchRows { get; set; }

        /// <summary>The password that encrypts the workbook.</summary>
        public string? Password { get; set; }
    }
}
