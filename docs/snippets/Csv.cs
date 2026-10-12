namespace Snippets.Csv
{
    internal sealed class ChangeRow
    {
        [ExcelColumn("file")]
        public string File { get; set; } = "";

        [ExcelColumn("lines_added")]
        public int LinesAdded { get; set; }
    }
}

namespace Snippets.Csv.Read
{
    #region csv-read
    using ExcelReader.Core.Parser;
    using ExcelReader.Core.Reader;
    #endregion

    internal static class Snippet
    {
        internal static void Run()
        {
            #region csv-read
            using var csv = Excel.FromCsvFile("report.csv");

            foreach (var row in csv.FirstSheet)
            {
                Console.WriteLine(row[0].GetString());
            }

            // Typed parsing works exactly like the Excel workbooks:
            foreach (var item in ExcelParser.FromAttributes<ChangeRow>().Parse(csv.FirstSheet))
            {
                Console.WriteLine($"{item.File}: +{item.LinesAdded}");
            }
            #endregion
        }
    }
}

namespace Snippets.Csv.ReadOptions
{
    internal static class Snippet
    {
        internal static void Run()
        {
            #region csv-read-options
            var options = new ExcelReaderOptions { Csv = new CsvReaderOptions { Delimiter = (byte)';' } };
            using var csv = Excel.FromCsvFile("relatorio.csv", options);
            #endregion
        }
    }
}

namespace Snippets.Csv.OpenByFormat
{
    internal static class Snippet
    {
        internal static void Run(string path, string secret)
        {
            #region csv-open-by-format
            ExcelFileFormat format = Path.GetExtension(path) is ".csv" ? ExcelFileFormat.Csv : ExcelFileFormat.Unknown;

            using IExcelWorkbook workbook = Excel.Open(path, format, new ExcelReaderOptions
            {
                Password = secret,                                                  // used by the Excel formats
                Csv = CsvReaderOptions.Default with { SniffDialect = true },        // used by the CSV format
            });
            #endregion
        }
    }
}

namespace Snippets.Csv.Sniff
{
    #region csv-sniff
    using ExcelReader.Core.Reader;
    using ExcelReader.Core.Reader.Csv;
    #endregion

    internal static class Snippet
    {
        internal static void Run()
        {
            #region csv-sniff
            CsvDialect dialect = CsvSniffer.DetectFile("export.csv");
            using var csv = Excel.FromCsvFile("export.csv", new ExcelReaderOptions { Csv = CsvReaderOptions.Default.WithDialect(dialect) });
            #endregion
        }
    }
}

namespace Snippets.Csv.SniffOnOpen
{
    internal static class Snippet
    {
        internal static void Run()
        {
            #region csv-sniff-on-open
            using var csv = Excel.FromCsvFile("export.csv", new ExcelReaderOptions { Csv = CsvReaderOptions.Default with { SniffDialect = true } });
            #endregion
        }
    }
}

namespace Snippets.Csv.Write
{
    #region csv-write
    using ExcelReader.Core.Writer;
    using ExcelReader.Core.Writer.Csv;
    #endregion

    internal static class Snippet
    {
        internal static void Run()
        {
            #region csv-write
            using var stream = File.Create("out.csv");
            using var workbook = CsvWorkbookWriter.Create(stream);
            using var writer = workbook.AddSheet("Sheet1");

            using (CsvRowWriter row = writer.StartRow())
            {
                row.Write("Name");
                row.Write("Total");
                row.Write("Created");
            }

            using (CsvRowWriter row = writer.StartRow())
            {
                row.Write("Q1");
                row.Write(42);
                row.Write(DateTime.UtcNow);
            }
            #endregion
        }
    }
}

namespace Snippets.Csv.WriteOptions
{
    internal static class Snippet
    {
        internal static void Run(Stream stream)
        {
            #region csv-write-options
            using var workbook = CsvWorkbookWriter.Create(stream, leaveOpen: false, new CsvWriterOptions
            {
                Delimiter = (byte)';',
                Encoding = Encoding.Latin1,          // fields are formatted as UTF-8, then transcoded on the way out
                WriteByteOrderMark = true,           // what makes Excel open a UTF-8 file as UTF-8
                NewLine = CsvNewLine.LineFeed,       // default is CRLF, as RFC 4180 specifies
            });
            #endregion
        }
    }
}
