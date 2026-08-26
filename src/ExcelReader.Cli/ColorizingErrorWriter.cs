using System.Text;
using Spectre.Console;

namespace ExcelReader.Cli
{
    /// <summary>
    /// Wraps standard error so <see cref="CliCommands.Execute"/>'s one-line failure message renders
    /// in red on an interactive terminal, and as the exact same plain text everywhere else.
    /// </summary>
    /// <remarks>
    /// <see cref="CliCommands"/> only ever calls <see cref="WriteLine(string?)"/> on the
    /// <c>TextWriter</c> it's given for errors, so that is the only member this class needs to give
    /// real behavior to - every other <see cref="TextWriter"/> member falls back to <paramref name="inner"/>
    /// unused. Kept out of <c>CliCommands.cs</c> deliberately: that file's whole point is a tested
    /// surface with no <c>Console</c>-shaped state, and this class exists only to decide, from
    /// <see cref="System.Console.IsErrorRedirected"/>, how a byte reaches a real terminal.
    /// </remarks>
    internal sealed class ColorizingErrorWriter(TextWriter inner) : TextWriter
    {
        public override Encoding Encoding => inner.Encoding;

        public override void WriteLine(string? value)
        {
            if (value is null)
            {
                inner.WriteLine();
                return;
            }

            if (System.Console.IsErrorRedirected)
            {
                inner.WriteLine(value);
                return;
            }

            ErrorConsole.Console.MarkupLine($"[red]{Markup.Escape(value)}[/]");
        }
    }
}
