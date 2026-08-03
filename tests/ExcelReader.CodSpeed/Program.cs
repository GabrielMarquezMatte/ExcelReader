using System.Globalization;

namespace ExcelReader.CodSpeed
{
    // Entry point of the CodSpeed harness.
    //
    // ExcelReader is a .NET library and CodSpeed has no .NET integration, so the benchmarks run
    // through the CodSpeed CLI's command harness: the CLI executes this program once per benchmark
    // and measures the process. Each invocation runs exactly one scenario, which keeps the
    // measurement attributable to a single code path.
    //
    // Usage:
    //   ExcelReader.CodSpeed <scenario> [--iterations N]
    //   ExcelReader.CodSpeed --list
    internal static class Program
    {
        public static async Task<int> Main(string[] args)
        {
            ArgumentNullException.ThrowIfNull(args);

            if (args.Length == 0)
            {
                WriteUsage();
                return 1;
            }

            switch (args[0])
            {
                case "--help" or "-h":
                    WriteUsage();
                    return 0;
                case "--list":
                    foreach (Scenario listed in ScenarioRegistry.All)
                    {
                        Console.WriteLine(listed.Name);
                    }
                    return 0;
                default:
                    break;
            }

            Scenario? scenario = ScenarioRegistry.Find(args[0]);
            if (scenario is null)
            {
                Console.Error.WriteLine($"Unknown scenario '{args[0]}'. Run with --list to see the available ones.");
                return 1;
            }

            if (!TryReadIterations(args, scenario.Iterations, out int iterations))
            {
                return 1;
            }

            long checksum = await scenario.RunAsync(iterations);

            // Printing the checksum keeps the workload observable, so nothing measured can be
            // optimized away as dead code.
            Console.WriteLine(string.Create(
                CultureInfo.InvariantCulture,
                $"{scenario.Name}: iterations={iterations} checksum={checksum}"));
            return 0;
        }

        private static bool TryReadIterations(string[] args, int defaultIterations, out int iterations)
        {
            iterations = defaultIterations;
            int index = 1;
            while (index < args.Length)
            {
                if (args[index] is not ("--iterations" or "-n"))
                {
                    Console.Error.WriteLine($"Unknown argument '{args[index]}'.");
                    return false;
                }

                if (index + 1 >= args.Length
                    || !int.TryParse(args[index + 1], NumberStyles.Integer, CultureInfo.InvariantCulture, out iterations)
                    || iterations < 1)
                {
                    Console.Error.WriteLine("--iterations requires a positive integer.");
                    return false;
                }

                index += 2;
            }
            return true;
        }

        private static void WriteUsage()
        {
            Console.WriteLine("Usage: ExcelReader.CodSpeed <scenario> [--iterations N]");
            Console.WriteLine("       ExcelReader.CodSpeed --list");
        }
    }
}
