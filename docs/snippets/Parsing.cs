namespace Snippets.Parsing
{
    internal sealed class ChangeRow
    {
        [ExcelColumn("file")]
        public string File { get; set; } = "";

        [ExcelColumn("lines_added")]
        public int LinesAdded { get; set; }
    }

    internal readonly ref struct ChangeRowRef
    {
        public ReadOnlySpan<byte> File { get; init; }
        public int LinesAdded { get; init; }
    }
}

namespace Snippets.Parsing.Attributes
{
    #region parsing-attributes
    using ExcelReader.Core.Parser;
    using ExcelReader.Core.Reader;
    #endregion

    #region parsing-attributes
    public sealed class ChangeRow
    {
        [ExcelColumn("file")]
        public string File { get; set; } = "";

        [ExcelColumn("lines_added")]
        public int LinesAdded { get; set; }
    }
    #endregion

    internal static class Snippet
    {
        internal static void Run()
        {
            #region parsing-attributes
            using var workbook = Excel.FromXlsxFile("changes.xlsx");
            var parser = ExcelParser.FromAttributes<ChangeRow>();

            foreach (var item in parser.Parse(workbook.FirstSheet))
            {
                Console.WriteLine($"{item.File}: +{item.LinesAdded}");
            }
            #endregion
        }
    }
}

namespace Snippets.Parsing.AttributesOpen
{
    internal static class Snippet
    {
        internal static void Run()
        {
            #region parsing-attributes-open
            using IExcelWorkbook workbook = Excel.Open("changes.xlsx"); // or .xlsb / .xls
            foreach (var item in ExcelParser.FromAttributes<ChangeRow>().Parse(workbook.FirstSheet)) { /* ... */ }
            #endregion
        }
    }
}

namespace Snippets.Parsing.Generated
{
    #region parsing-generated
    using ExcelReader.Core.Parser;
    using ExcelReader.Core.Reader;
    using ExcelReader.Core.Writer;
    using ExcelReader.Core.Writer.Xlsx;
    #endregion

    #region parsing-generated
    [ExcelSerializable]                       // the model must be declared partial
    public partial class ChangeRow
    {
        [ExcelColumn("file")]
        public string File { get; set; } = "";

        [ExcelColumn("lines_added")]
        public int LinesAdded { get; set; }
    }
    #endregion

    internal static class Snippet
    {
        internal static async Task Run()
        {
            #region parsing-generated
            using var workbook = Excel.FromXlsxFile("changes.xlsx");
            foreach (var item in ExcelParser.Generated<ChangeRow>().Parse(workbook.FirstSheet))
            {
                Console.WriteLine($"{item.File}: +{item.LinesAdded}");
            }

            var changes = new[] { new ChangeRow { File = "README.md", LinesAdded = 12 } };
            var stream = File.Create("changes.xlsx");
            await using (stream.ConfigureAwait(false))
            {
                await using var writer = XlsxWorkbookWriter.Create(stream).ConfigureAwait(false);   // or XlsbWorkbookWriter / XlsWorkbookWriter / CsvWorkbookWriter
                await using (var sheet = writer.AddSheet("Changes").ConfigureAwait(false))
                {
                    await sheet.WriteRecordsAsync(changes, ExcelRecordLayout.Generated<ChangeRow>()).ConfigureAwait(false);
                }
            }
            #endregion
        }
    }
}

namespace Snippets.Parsing.Fluent
{
    #region parsing-fluent
    using ExcelReader.Core.Parser;
    using ExcelReader.Core.Reader;
    #endregion

    #region parsing-fluent
    public sealed class ChangeRow
    {
        public string File { get; set; } = "";
        public int LinesAdded { get; set; }
    }
    #endregion

    internal static class Snippet
    {
        internal static void Run()
        {
            #region parsing-fluent
            var parser = ExcelParser.Build<ChangeRow>(builder => builder
                .Factory(() => new ChangeRow())
                .Property(["file"], ExcelCellReaders.String, (ref ChangeRow r, string v) => r.File = v)
                .Property(["lines_added"], ExcelCellReaders.Parsable, (ref ChangeRow r, int v) => r.LinesAdded = v));

            using var workbook = Excel.FromXlsxFile("changes.xlsx");
            foreach (var item in parser.Parse(workbook.FirstSheet))
            {
                Console.WriteLine($"{item.File}: +{item.LinesAdded}");
            }
            #endregion
        }
    }
}

namespace Snippets.Parsing.FluentIndex
{
    internal static class Snippet
    {
        internal static void Run()
        {
            #region parsing-fluent-index
            var parser = ExcelParser.Build<ChangeRow>(builder => builder
                .Factory(() => new ChangeRow())
                .PropertyAt(0, ExcelCellReaders.String, (ref ChangeRow r, string v) => r.File = v)
                .PropertyAt(1, ExcelCellReaders.Parsable, (ref ChangeRow r, int v) => r.LinesAdded = v));
            #endregion
        }
    }
}

namespace Snippets.Parsing.FluentFallback
{
    internal static class Snippet
    {
        internal static void Run()
        {
            #region parsing-fluent-fallback
            var parser = ExcelParser.BuildWithAttributeFallback<ChangeRow>(builder => builder
                .Property(["file"], ExcelCellReaders.String, (ref ChangeRow r, string v) => r.File = v.ToUpperInvariant()));
            #endregion
        }
    }
}

namespace Snippets.Parsing.Options
{
    #region parsing-options
    using System.Globalization;
    using ExcelReader.Core.Parser;
    #endregion

    internal static class Snippet
    {
        internal static void Run()
        {
            #region parsing-options
            var options = new ExcelParserOptions
            {
                HeaderRow = 1,                                   // 1-based row holding the headers
                ColumnNameComparer = StringComparer.OrdinalIgnoreCase,
                HeaderNormalization = HeaderNormalization.Trim | HeaderNormalization.CollapseSpaces,
                Culture = CultureInfo.GetCultureInfo("pt-BR"),   // parse "1.234,56" as 1234.56m
            };

            var parser = ExcelParser.FromAttributes<ChangeRow>(options);
            #endregion
        }
    }
}

namespace Snippets.Parsing.Required
{
    #region parsing-required
    public sealed class Order
    {
        [ExcelRequired]
        public int Id { get; set; }

        [ExcelRequired(AllowEmpty = true)]   // column must exist; blank cells allowed
        public string? Note { get; set; }
    }
    #endregion
}

namespace Snippets.Parsing.Converter
{
    #region parsing-converter
    using System.Globalization;
    using ExcelReader.Core.Parser;
    using ExcelReader.Core.Reader;
    #endregion

    #region parsing-converter
    public sealed class BrlMoneyConverter : IExcelCellConverter<decimal>
    {
        public bool TryConvert(in Cell cell, in ExcelRowContext context, out decimal value)
        {
            string text = cell.GetString().Replace("R$", "", StringComparison.Ordinal).Trim();
            return decimal.TryParse(text, NumberStyles.Currency, CultureInfo.GetCultureInfo("pt-BR"), out value);
        }
    }

    public sealed class Invoice
    {
        [ExcelConverter(typeof(BrlMoneyConverter))]
        public decimal Total { get; set; }
    }
    #endregion
}

namespace Snippets.Parsing.RefStruct
{
    #region parsing-ref-struct
    using System.Text;
    using ExcelReader.Core.Parser;
    using ExcelReader.Core.Reader;
    #endregion

    #region parsing-ref-struct
    public readonly ref struct ChangeRowRef
    {
        public ReadOnlySpan<byte> File { get; init; }   // zero-copy — aliases the reader's row buffer
        public int LinesAdded { get; init; }
    }
    #endregion

    internal static class Snippet
    {
        internal static void Run()
        {
            #region parsing-ref-struct
            using var workbook = Excel.FromXlsxFile("changes.xlsx");

            foreach (ChangeRowRef item in ExcelParser.FromAttributes<ChangeRowRef>().Parse(workbook.FirstSheet))
            {
                Console.WriteLine($"{Encoding.UTF8.GetString(item.File)}: +{item.LinesAdded}");
            }
            #endregion
        }
    }
}

namespace Snippets.Parsing.RefStructAsync
{
    internal static class Snippet
    {
        internal static async Task Run()
        {
            #region parsing-ref-struct-async
            var workbook = await Excel.FromXlsxFileAsync("changes.xlsx").ConfigureAwait(false);
            await using (workbook.ConfigureAwait(false))
            {
                await foreach (ChangeRowRef item in ExcelParser.FromAttributes<ChangeRowRef>().Parse(workbook.FirstSheet).ConfigureAwait(false))
                {
                    Console.WriteLine($"{Encoding.UTF8.GetString(item.File)}: +{item.LinesAdded}");
                }
            }
            #endregion
        }
    }
}
