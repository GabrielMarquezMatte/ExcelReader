using System.Diagnostics.CodeAnalysis;
using ConsoleAppFramework;

namespace ExcelReader.Cli
{
    [ExcludeFromCodeCoverage]
    internal static class Program
    {
        internal static int Main(string[] args)
        {
            ConsoleApp.ConsoleAppBuilder app = ConsoleApp.Create();
            app.Add<Commands>();
            app.Run(args);
            return Environment.ExitCode;
        }
    }
}
