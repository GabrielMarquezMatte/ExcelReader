using BenchmarkDotNet.Running;

namespace ExcelReader.Benchmarks
{
    internal static class Program
    {
        // Run all benchmarks, or filter: `dotnet run -c Release -- --filter *Write*`
        public static void Main(string[] args)
        {
            BenchmarkSwitcher.FromAssembly(typeof(Program).Assembly).Run(args);
        }
    }
}
