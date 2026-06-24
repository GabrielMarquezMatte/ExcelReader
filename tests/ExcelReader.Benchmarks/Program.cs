using BenchmarkDotNet.Configs;
using BenchmarkDotNet.Exporters;
using BenchmarkDotNet.Exporters.Json;
using BenchmarkDotNet.Running;

namespace ExcelReader.Benchmarks
{
    internal static class Program
    {
        private static readonly IConfig Config = ManualConfig.Create(DefaultConfig.Instance)
            .AddExporter(MarkdownExporter.GitHub)
            .AddExporter(JsonExporter.Full);

        // Run all benchmarks, or filter: `dotnet run -c Release -- --filter *Write*`
        public static void Main(string[] args)
        {
            BenchmarkSwitcher.FromAssembly(typeof(Program).Assembly).Run(args, Config);
        }
    }
}
