using Spectre.Console;

namespace ExcelReader.Cli
{
    /// <summary>
    /// A Spectre.Console rendering surface permanently pointed at standard error instead of the
    /// default standard output.
    /// </summary>
    /// <remarks>
    /// <see cref="AnsiConsole"/>'s own instance writes to stdout - correct for <c>sheets</c>/<c>schema</c>'s
    /// tables, since a table there *is* the command's result. It is wrong for anything that isn't a
    /// command's actual result - <c>convert</c>'s progress spinner, and every command's error text -
    /// because <c>convert</c>'s stdout may be carrying piped CSV or a binary workbook that must stay
    /// byte-exact; writing so much as one spinner frame to it would corrupt the stream.
    /// </remarks>
    internal static class ErrorConsole
    {
        internal static readonly IAnsiConsole Console = AnsiConsole.Create(new AnsiConsoleSettings
        {
            Out = new AnsiConsoleOutput(System.Console.Error),
        });
    }
}
