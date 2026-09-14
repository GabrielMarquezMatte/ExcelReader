using System.Text;

namespace ExcelReader.Cli
{
    // Wraps standard error so CliCommands.Execute's one-line failure message renders
    // in red on an interactive terminal, and as the exact same plain text everywhere else.
    //
    // CliCommands only ever calls WriteLine(string?) on the
    // TextWriter it's given for errors, so that is the only member this class needs to give
    // real behavior to - every other TextWriter member falls back to inner
    // unused. Kept out of CliCommands.cs deliberately: that file's whole point is a tested
    // surface with no Console-shaped state, and this class exists only to decide, from
    // System.Console.IsErrorRedirected, how a byte reaches a real terminal. A raw ANSI
    // escape is enough for one red line - no need for Spectre.Console's markup renderer here.
    internal sealed class ColorizingErrorWriter(TextWriter inner) : TextWriter
    {
        private const string Red = "\u001b[31m";
        private const string Reset = "\u001b[0m";

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

            inner.WriteLine($"{Red}{value}{Reset}");
        }
    }
}
