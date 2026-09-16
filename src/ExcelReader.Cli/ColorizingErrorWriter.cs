using System.Diagnostics.CodeAnalysis;
using System.Text;

namespace ExcelReader.Cli
{
    [ExcludeFromCodeCoverage]
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
