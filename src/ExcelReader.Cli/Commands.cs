using System.Globalization;
using ConsoleAppFramework;
using Spectre.Console;

namespace ExcelReader.Cli
{
    /// <summary>
    /// The command surface ConsoleAppFramework generates parsing, routing and help from. Every method
    /// is a one-line-or-two adapter: the real work lives in <see cref="CliCommands"/>, which takes its
    /// writers as parameters so the tests never touch ConsoleAppFramework's static output hooks or
    /// Spectre.Console's rendering.
    /// </summary>
    /// <remarks>
    /// Every command's interactive/plain split follows the same rule: when the relevant stream (stdout
    /// for <c>sheets</c>/<c>schema</c>'s result, stderr for <c>convert</c>'s progress) is redirected -
    /// piped, written to a file, or otherwise not a real terminal - nothing here changes: the same
    /// tab-separated text a script already parses, unchanged since before Spectre.Console existed.
    /// Only on an actual terminal does Spectre.Console's table/spinner render instead. This is not
    /// re-tested here on purpose - it's the ConsoleAppFramework precedent this file already followed
    /// (see the type-level remarks on that split): rendering is glue, not logic.
    /// </remarks>
    internal sealed class Commands
    {
        /// <summary>Lists every sheet in a workbook, as "index[TAB]name".</summary>
        /// <param name="path">Path to the workbook (.xlsx, .xlsb, .xls or .csv).</param>
        /// <param name="password">Password for an encrypted .xlsx/.xlsb workbook.</param>
        [Command("sheets")]
        public int Sheets([Argument] string path, string? password = null)
        {
            using TextWriter stderr = new ColorizingErrorWriter(Console.Error);
            if (Console.IsOutputRedirected)
            {
                return CliCommands.Sheets(path, Console.Out, stderr, password);
            }

            Table table = new Table().AddColumn("Index").AddColumn("Sheet");
            int code = CliCommands.Sheets(path, (index, name) =>
                table.AddRow(index.ToString(CultureInfo.InvariantCulture), Markup.Escape(name)), stderr, password);
            if (code == 0)
            {
                AnsiConsole.Write(table);
            }
            return code;
        }

        /// <summary>Writes a sheet out to another spreadsheet format, to a file or to standard output.</summary>
        /// <param name="path">Path to the workbook (.xlsx, .xlsb, .xls or .csv).</param>
        /// <param name="sheet">-s, Sheet to convert, by name or zero-based index. Defaults to the first.</param>
        /// <param name="output">-o, File to write to. Writes to standard output when omitted.</param>
        /// <param name="format">-f, Output format: xlsx, xlsb, xls or csv. Defaults to --output's extension, or csv when writing to standard output.</param>
        /// <param name="delimiter">-d, CSV field separator. Ignored for every other format. Defaults to a comma.</param>
        /// <param name="password">Password for an encrypted .xlsx/.xlsb source workbook.</param>
        [Command("convert")]
        public int Convert([Argument] string path, string? sheet = null, string? output = null, string? format = null, char? delimiter = null, string? password = null)
        {
            // ConsoleAppFramework's source generator mis-emits a `char` parameter whose default
            // literal is itself a comma (the codegen that renders parameter defaults splits on ','),
            // so the default lives here instead of in the signature.
            using Stream stdout = Console.OpenStandardOutput();
            using TextWriter stderr = new ColorizingErrorWriter(Console.Error);

            // Stdout may be carrying the converted bytes themselves (piped CSV, or a binary workbook
            // written with no --output) - a spinner frame on a non-terminal stderr is one thing
            // scripts don't parse, but rendering anywhere near stdout here would be a real corruption
            // risk, so the progress UI is gated on stderr alone, never on stdout's redirected state.
            if (Console.IsErrorRedirected)
            {
                return CliCommands.Convert(path, sheet, output, format, delimiter ?? ',', stdout, stderr, onProgress: null, password);
            }

            return ErrorConsole.Console.Status().Start("Converting...", ctx =>
                CliCommands.Convert(path, sheet, output, format, delimiter ?? ',', stdout, stderr, rowsWritten =>
                    ctx.Status($"Converting... {rowsWritten.ToString("N0", CultureInfo.InvariantCulture)} rows written"), password));
        }

        /// <summary>Prints the inferred column schema, as "index[TAB]name[TAB]type", with a trailing ? for nullable columns.</summary>
        /// <param name="path">Path to the workbook (.xlsx, .xlsb, .xls or .csv).</param>
        /// <param name="sheet">-s, Sheet to inspect, by name or zero-based index. Defaults to the first.</param>
        /// <param name="headerRow">1-based row to take column names from; 0 means the sheet has no header.</param>
        /// <param name="sampleSize">How many rows after the header to sample.</param>
        /// <param name="password">Password for an encrypted .xlsx/.xlsb workbook.</param>
        [Command("schema")]
        public int Schema([Argument] string path, string? sheet = null, int headerRow = 1, int sampleSize = 100, string? password = null)
        {
            using TextWriter stderr = new ColorizingErrorWriter(Console.Error);
            if (Console.IsOutputRedirected)
            {
                return CliCommands.Schema(path, sheet, headerRow, sampleSize, Console.Out, stderr, password);
            }

            Table table = new Table().AddColumn("Index").AddColumn("Name").AddColumn("Type").AddColumn("Nullable");
            int code = CliCommands.Schema(path, sheet, headerRow, sampleSize, column =>
                table.AddRow(
                    column.Index.ToString(CultureInfo.InvariantCulture),
                    Markup.Escape(column.Name ?? string.Empty),
                    column.Type.ToString(),
                    column.IsNullable ? "yes" : "no"),
                stderr, password);
            if (code == 0)
            {
                AnsiConsole.Write(table);
            }
            return code;
        }
    }
}
