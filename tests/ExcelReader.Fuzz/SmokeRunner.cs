namespace ExcelReader.Fuzz
{
    /// <summary>
    /// Runs every fuzz target over a corpus plus deterministic mutations of it, with no native
    /// fuzzing engine involved.
    /// </summary>
    /// <remarks>
    /// Two jobs. First, it validates the harnesses and <see cref="FuzzOracle"/> themselves — a fuzz
    /// suite whose oracle is wrong finds nothing, and that failure is silent. Second, it is a
    /// dumb-but-free fuzzer that any CI run can afford, so the readers get some adversarial input on
    /// every commit rather than only on the nightly libFuzzer job. It is strictly weaker than
    /// coverage-guided fuzzing: mutations are blind, so it explores shallowly.
    /// </remarks>
    internal static class SmokeRunner
    {
        internal static int Run(string corpusDirectory, int mutationsPerInput, int seed)
        {
            FuzzOracle.SelfCheck();

            string[] files = Directory.Exists(corpusDirectory)
                ? Directory.GetFiles(corpusDirectory)
                : [];
            if (files.Length == 0)
            {
                Console.Error.WriteLine($"no corpus files in {corpusDirectory}");
                return 1;
            }

            var random = new Random(seed);
            int failures = 0;
            int executed = 0;

            foreach (string file in files.Order(StringComparer.Ordinal))
            {
                byte[] original = File.ReadAllBytes(file);
                foreach ((string name, Action<ReadOnlySpan<byte>> target) in Program.AllTargets)
                {
                    failures += RunOne(name, target, original, file, "as-is", ref executed);
                    for (int i = 0; i < mutationsPerInput; i++)
                    {
                        byte[] mutated = Mutate(original, random);
                        failures += RunOne(name, target, mutated, file, $"mutation #{i}", ref executed);
                    }
                }
            }

            Console.WriteLine($"executed {executed} case(s) across {files.Length} corpus file(s); {failures} failure(s)");
            return failures == 0 ? 0 : 1;
        }

        private static int RunOne(
            string targetName,
            Action<ReadOnlySpan<byte>> target,
            byte[] input,
            string sourceFile,
            string what,
            ref int executed)
        {
            executed++;
            try
            {
                target(input);
                return 0;
            }
            catch (Exception ex)
            {
                // FuzzOracle already let the sanctioned exceptions through inside the harness, so
                // anything arriving here is by definition unexpected.
                Console.Error.WriteLine($"FAIL target={targetName} source={Path.GetFileName(sourceFile)} input={what}");
                Console.Error.WriteLine($"  {ex.GetType().FullName}: {ex.Message}");
                Console.Error.WriteLine(ex.StackTrace);
                string dump = Path.Combine(Path.GetTempPath(), $"fuzz-fail-{targetName}-{Guid.NewGuid():N}.bin");
                File.WriteAllBytes(dump, input);
                Console.Error.WriteLine($"  input written to {dump}");
                return 1;
            }
        }

        // Blind byte-level mutations. Truncation matters most here: it is how a reader is made to
        // meet an offset or length field that points past the end of the data.
        private static byte[] Mutate(byte[] original, Random random)
        {
            byte[] copy;
            switch (random.Next(4))
            {
                case 0: // truncate
                    copy = original[..random.Next(0, original.Length + 1)];
                    break;
                case 1: // flip a handful of bits
                    copy = [.. original];
                    for (int i = 0; i < 8 && copy.Length > 0; i++)
                    {
                        int at = random.Next(copy.Length);
                        copy[at] ^= (byte)(1 << random.Next(8));
                    }
                    break;
                case 2: // overwrite a run with a repeated byte (drives length/count fields to extremes)
                    copy = [.. original];
                    if (copy.Length > 0)
                    {
                        int start = random.Next(copy.Length);
                        int length = Math.Min(copy.Length - start, random.Next(1, 17));
                        byte value = (byte)random.Next(256);
                        copy.AsSpan(start, length).Fill(value);
                    }
                    break;
                default: // splice a chunk over another position
                    copy = [.. original];
                    if (copy.Length > 4)
                    {
                        int length = random.Next(1, Math.Min(64, copy.Length));
                        int from = random.Next(copy.Length - length + 1);
                        int to = random.Next(copy.Length - length + 1);
                        copy.AsSpan(from, length).CopyTo(copy.AsSpan(to));
                    }
                    break;
            }
            return copy;
        }
    }
}
