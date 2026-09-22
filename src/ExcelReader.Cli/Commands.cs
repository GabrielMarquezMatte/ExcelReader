using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using ConsoleAppFramework;

namespace ExcelReader.Cli
{
    [ExcludeFromCodeCoverage]
    internal sealed class Commands
    {
        [Command("sheets")]
        public int Sheets([Argument] string path, string? password = null)
        {
            using TextWriter stderr = new ColorizingErrorWriter(Console.Error);
            return CliCommands.Sheets(path, Console.Out, stderr, password);
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

            // ponytail: progress is a carriage-returned line on stderr rather than a spinner widget. A
            int code = CliCommands.Convert(path, sheet, output, format, delimiter ?? ',', stdout, stderr, rowsWritten =>
                Console.Error.Write($"\rConverting... {rowsWritten.ToString("N0", CultureInfo.InvariantCulture)} rows written"), password);
            Console.Error.WriteLine();
            return code;
        }

        [Command("schema")]
        public int Schema([Argument] string path, string? sheet = null, int headerRow = 1, int sampleSize = 100, string? password = null)
        {
            using TextWriter stderr = new ColorizingErrorWriter(Console.Error);
            return CliCommands.Schema(path, sheet, headerRow, sampleSize, Console.Out, stderr, password);
        }
    }
}
