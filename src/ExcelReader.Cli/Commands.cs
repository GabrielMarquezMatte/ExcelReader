using ConsoleAppFramework;

namespace ExcelReader.Cli
{
    /// <summary>
    /// The command surface ConsoleAppFramework generates parsing, routing and help from. Every method
    /// is a one-line adapter: the real work lives in <see cref="CliCommands"/>, which takes its
    /// writers as parameters so the tests never touch ConsoleAppFramework's static output hooks.
    /// </summary>
    internal sealed class Commands
    {
        /// <summary>Lists every sheet in a workbook, as "index[TAB]name".</summary>
        /// <param name="path">Path to the workbook (.xlsx, .xlsb, .xls or .csv).</param>
        [Command("sheets")]
        public int Sheets([Argument] string path)
        {
            return CliCommands.Sheets(path, Console.Out, Console.Error);
        }

        /// <summary>Writes a sheet out to another spreadsheet format, to a file or to standard output.</summary>
        /// <param name="path">Path to the workbook (.xlsx, .xlsb, .xls or .csv).</param>
        /// <param name="sheet">-s, Sheet to convert, by name or zero-based index. Defaults to the first.</param>
        /// <param name="output">-o, File to write to. Writes to standard output when omitted.</param>
        /// <param name="format">-f, Output format: xlsx, xlsb, xls or csv. Defaults to --output's extension, or csv when writing to standard output.</param>
        /// <param name="delimiter">-d, CSV field separator. Ignored for every other format. Defaults to a comma.</param>
        [Command("convert")]
        public int Convert([Argument] string path, string? sheet = null, string? output = null, string? format = null, char? delimiter = null)
        {
            // ConsoleAppFramework's source generator mis-emits a `char` parameter whose default
            // literal is itself a comma (the codegen that renders parameter defaults splits on ','),
            // so the default lives here instead of in the signature.
            using Stream stdout = Console.OpenStandardOutput();
            return CliCommands.Convert(path, sheet, output, format, delimiter ?? ',', stdout, Console.Error);
        }

        /// <summary>Prints the inferred column schema, as "index[TAB]name[TAB]type", with a trailing ? for nullable columns.</summary>
        /// <param name="path">Path to the workbook (.xlsx, .xlsb, .xls or .csv).</param>
        /// <param name="sheet">-s, Sheet to inspect, by name or zero-based index. Defaults to the first.</param>
        /// <param name="headerRow">1-based row to take column names from; 0 means the sheet has no header.</param>
        /// <param name="sampleSize">How many rows after the header to sample.</param>
        [Command("schema")]
        public int Schema([Argument] string path, string? sheet = null, int headerRow = 1, int sampleSize = 100)
        {
            return CliCommands.Schema(path, sheet, headerRow, sampleSize, Console.Out, Console.Error);
        }
    }
}
