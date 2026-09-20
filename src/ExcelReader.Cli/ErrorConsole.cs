using System.Diagnostics.CodeAnalysis;
using Spectre.Console;

namespace ExcelReader.Cli
{
    [ExcludeFromCodeCoverage]
    internal static class ErrorConsole
    {
        internal static readonly IAnsiConsole Console = AnsiConsole.Create(new AnsiConsoleSettings
        {
            Out = new AnsiConsoleOutput(System.Console.Error),
        });
    }
}
