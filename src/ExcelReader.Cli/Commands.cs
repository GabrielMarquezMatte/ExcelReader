using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Text;
using ConsoleAppFramework;

namespace ExcelReader.Cli
{
    [ExcludeFromCodeCoverage]
    internal sealed class Commands
    {
        [Command("sheets")]
        public int Sheets([Argument] string path, string? password = null)
        {
            using TextWriter stdout = StandardOutput();
            using TextWriter stderr = new ColorizingErrorWriter(Console.Error);
            return CliCommands.Sheets(path, stdout, stderr, password);
        }

        /// <param name="sheet">-s, Sheet name or 0-based index.</param>
        /// <param name="output">-o, Output file; its extension picks the format.</param>
        /// <param name="format">-f, Output format: xlsx, xlsb, xls or csv.</param>
        /// <param name="delimiter">-d, CSV output delimiter.</param>
        [Command("convert")]
        public int Convert([Argument] string path, string? sheet = null, string? output = null, string? format = null, char? delimiter = null, char? inputDelimiter = null, string? password = null)
        {
            using Stream stdout = Console.OpenStandardOutput();
            using TextWriter stderr = new ColorizingErrorWriter(Console.Error);

            if (Console.IsErrorRedirected)
            {
                return CliCommands.Convert(path, sheet, output, format, delimiter ?? ',', stdout, stderr, onProgress: null, password, inputDelimiter);
            }

            // ponytail: progress is a carriage-returned line on stderr rather than a spinner widget.
            int code = CliCommands.Convert(path, sheet, output, format, delimiter ?? ',', stdout, stderr, rowsWritten =>
                Console.Error.Write($"\rConverting... {rowsWritten.ToString("N0", CultureInfo.InvariantCulture)} rows written"), password, inputDelimiter);
            Console.Error.WriteLine();
            return code;
        }

        /// <param name="sheet">-s, Sheet name or 0-based index.</param>
        [Command("schema")]
        public int Schema([Argument] string path, string? sheet = null, int headerRow = 1, int sampleSize = 100, char? inputDelimiter = null, string? password = null)
        {
            using TextWriter stdout = StandardOutput();
            using TextWriter stderr = new ColorizingErrorWriter(Console.Error);
            return CliCommands.Schema(path, sheet, headerRow, sampleSize, stdout, stderr, password, inputDelimiter);
        }

        /// <param name="sheet">-s, Sheet name or 0-based index.</param>
        [Command("preview")]
        public int Preview([Argument] string path, string? sheet = null, int rows = 10, int maxColumns = 8, int maxWidth = 32, int headerRow = 1, bool fast = false, string? columns = null, char? inputDelimiter = null, string? password = null)
        {
            using TextWriter stdout = StandardOutput();
            using TextWriter stderr = new ColorizingErrorWriter(Console.Error);
            return CliCommands.Preview(path, sheet, headerRow, rows, maxColumns, maxWidth, fast, columns, stdout, stderr, password, inputDelimiter);
        }

        // UTF-16 on a Windows console makes .NET use WriteConsoleW and leave the shell's code page alone.
        // Not with stderr redirected: Console.Error would write UTF-16 bytes into that file.
        private static TextWriter StandardOutput()
        {
            if (Console.IsOutputRedirected || !OperatingSystem.IsWindows())
            {
                return new StreamWriter(Console.OpenStandardOutput(), new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));
            }
            if (!Console.IsErrorRedirected)
            {
                Console.OutputEncoding = Encoding.Unicode;
            }
            return Console.Out;
        }
    }
}
