using System.Globalization;
using SharpFuzz;

namespace ExcelReader.Fuzz
{
    /// <summary>
    /// Entry point for the fuzz harnesses, in three modes: <c>seeds</c> writes a starting corpus,
    /// <c>check</c> runs the targets without a fuzzing engine, and naming a target runs it under
    /// libFuzzer.
    /// </summary>
    /// <remarks>
    /// libFuzzer mode requires ExcelReader.Core to be instrumented with the <c>sharpfuzz</c> CLI and
    /// the process to be launched by <c>libfuzzer-dotnet</c> — see README.md in this directory.
    /// </remarks>
    internal static class Program
    {
        internal static readonly Dictionary<string, Action<ReadOnlySpan<byte>>> AllTargets =
            new(StringComparer.OrdinalIgnoreCase)
            {
                ["xlsx"] = Harnesses.Xlsx,
                ["xlsx-memory"] = Harnesses.XlsxMemory,
                ["xlsb"] = Harnesses.Xlsb,
                ["xlsb-memory"] = Harnesses.XlsbMemory,
                ["xls"] = Harnesses.Xls,
                ["csv"] = Harnesses.Csv,
                ["csv-sniff"] = Harnesses.CsvSniff,
            };

        private static async Task<int> Main(string[] args)
        {
            if (args.Length >= 2 && string.Equals(args[0], "seeds", StringComparison.OrdinalIgnoreCase))
            {
                await SeedCorpus.GenerateAsync(args[1]);
                return 0;
            }

            if (args.Length >= 2 && string.Equals(args[0], "check", StringComparison.OrdinalIgnoreCase))
            {
                int mutations = args.Length >= 3 ? int.Parse(args[2], CultureInfo.InvariantCulture) : 64;
                int seed = args.Length >= 4 ? int.Parse(args[3], CultureInfo.InvariantCulture) : 1;
                return SmokeRunner.Run(args[1], mutations, seed);
            }

            // libfuzzer-dotnet forwards its own flags too, so match on the first argument only.
            string? name = args.Length > 0 ? args[0] : Environment.GetEnvironmentVariable("FUZZ_TARGET");
            if (name is null || !AllTargets.TryGetValue(name, out Action<ReadOnlySpan<byte>>? target))
            {
                Console.Error.WriteLine("usage:");
                Console.Error.WriteLine("  ExcelReader.Fuzz seeds <dir>                     write a starting corpus");
                Console.Error.WriteLine("  ExcelReader.Fuzz check <dir> [mutations] [seed]  run every target, no engine");
                Console.Error.WriteLine($"  ExcelReader.Fuzz <{string.Join('|', AllTargets.Keys)}>   run under libFuzzer");
                return 1;
            }

            Fuzzer.LibFuzzer.Run(span => target(span));
            return 0;
        }
    }
}
