using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using ConsoleAppFramework;
using Spectre.Console;

namespace ExcelReader.Cli
{
    [ExcludeFromCodeCoverage]
    internal sealed class Commands
    {
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

        [Command("convert")]
        public int Convert([Argument] string path, string? sheet = null, string? output = null, string? format = null, char? delimiter = null, string? password = null)
        {
            using Stream stdout = Console.OpenStandardOutput();
            using TextWriter stderr = new ColorizingErrorWriter(Console.Error);

            if (Console.IsErrorRedirected)
            {
                return CliCommands.Convert(path, sheet, output, format, delimiter ?? ',', stdout, stderr, onProgress: null, password);
            }

            return ErrorConsole.Console.Status().Start("Converting...", ctx =>
                CliCommands.Convert(path, sheet, output, format, delimiter ?? ',', stdout, stderr, rowsWritten =>
                    ctx.Status($"Converting... {rowsWritten.ToString("N0", CultureInfo.InvariantCulture)} rows written"), password));
        }

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
