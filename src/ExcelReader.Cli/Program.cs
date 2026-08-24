using ConsoleAppFramework;

namespace ExcelReader.Cli
{
    // An explicit entry-point class rather than top-level statements: ExcelReader.Tests is itself an
    // Exe, and a generated global-namespace Program would collide with its test host's entry point
    // once this project is referenced.
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
