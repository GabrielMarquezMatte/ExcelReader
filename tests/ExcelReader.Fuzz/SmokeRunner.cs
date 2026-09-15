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
            Harnesses.AssertSameRowsSelfCheck();
            VerifyEncryptedSeedReachesRealCode(corpusDirectory);
            VerifyDifferentialSeedsReachBothReaders(corpusDirectory);

            // Recursive: the corpus is laid out one directory per target (see SeedCorpus), but this
            // runner deliberately drives EVERY target over EVERY file. That cross-format pass is the
            // point of the Mutate splice below — it moves BIFF records into ZIP containers and back,
            // which is exactly the material a per-target libFuzzer corpus never produces.
            string[] files = Directory.Exists(corpusDirectory)
                ? [.. Directory.GetFiles(corpusDirectory, "*", SearchOption.AllDirectories).Order(StringComparer.Ordinal)]
                : [];
            if (files.Length == 0)
            {
                Console.Error.WriteLine($"no corpus files in {corpusDirectory}");
                return 1;
            }

            var random = new Random(seed);
            int failures = 0;
            int executed = 0;

            // Kept alongside the per-file mutation loop below so a mutation can splice a chunk from a
            // *different* corpus file — a shape blind single-file mutation can never produce, useful
            // here because it moves e.g. BIFF records from one container into another.
            byte[][] corpus = [.. files.Select(File.ReadAllBytes)];

            for (int f = 0; f < files.Length; f++)
            {
                byte[] original = corpus[f];
                foreach ((string name, Action<ReadOnlySpan<byte>> target) in Program.AllTargets)
                {
                    failures += RunOne(name, target, original, files[f], "as-is", ref executed);
                    for (int i = 0; i < mutationsPerInput; i++)
                    {
                        byte[] mutated = Mutate(original, corpus, random);
                        failures += RunOne(name, target, mutated, files[f], $"mutation #{i}", ref executed);
                    }
                }
            }

            Console.WriteLine($"executed {executed} case(s) across {files.Length} corpus file(s); {failures} failure(s)");
            return failures == 0 ? 0 : 1;
        }

        // The permanent regression guard for Critical 2 in the final review: the "encrypted" target
        // must reach AgileKeyDerivation/DecryptedPackageStream/PackageIntegrity, not dead-end on a
        // resource limit before ever touching them. That failure mode is invisible from failures==0
        // alone (a limit rejection and a genuine malformed-input rejection look identical to
        // FuzzOracle), so it's checked directly: the unmutated seed must actually open and yield rows.
        // Skips silently when the corpus directory has no such file (e.g. a scratch/partial corpus
        // used for a one-off repro) rather than failing an unrelated `check` run.
        private static void VerifyEncryptedSeedReachesRealCode(string corpusDirectory)
        {
            if (!Directory.Exists(corpusDirectory))
            {
                return;
            }

            // Both the committed regression seed and every generated one (SeedCorpus writes these
            // under the encryption password Harnesses opens with). A generated seed written under the
            // wrong password would be inert in exactly the same silent way.
            string[] seeds =
            [
                .. Directory.GetFiles(corpusDirectory, "encrypted-agile-seed.bin", SearchOption.AllDirectories),
                .. Directory.GetFiles(corpusDirectory, "seed-encrypted*.bin", SearchOption.AllDirectories),
            ];
            foreach (string seedPath in seeds.Order(StringComparer.Ordinal))
            {
                VerifyOneEncryptedSeed(seedPath);
            }
        }

        private static void VerifyDifferentialSeedsReachBothReaders(string corpusDirectory)
        {
            if (!Directory.Exists(corpusDirectory))
            {
                return;
            }

            foreach (string seedPath in Directory.GetFiles(corpusDirectory, "seed-xlsx*.bin", SearchOption.AllDirectories).Order(StringComparer.Ordinal))
            {
                VerifyOneDifferentialSeed(seedPath, xlsb: false);
            }
            foreach (string seedPath in Directory.GetFiles(corpusDirectory, "seed-xlsb*.bin", SearchOption.AllDirectories).Order(StringComparer.Ordinal))
            {
                VerifyOneDifferentialSeed(seedPath, xlsb: true);
            }
        }

        private static void VerifyOneDifferentialSeed(string seedPath, bool xlsb)
        {
            string name = Path.GetFileName(seedPath);
            int rows;
            try
            {
                rows = Harnesses.OpenDifferentialSeedForSelfCheck(File.ReadAllBytes(seedPath), xlsb);
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException(
                    $"the unmutated {name} no longer reads identically through both the stream and " +
                    "memory container readers - the matching differential target would compare two " +
                    $"rejections instead of two row sets, making it inert: {ex.GetType().FullName}: {ex.Message}",
                    ex);
            }
            if (rows == 0)
            {
                throw new InvalidOperationException(
                    $"the unmutated {name} opened through both readers but yielded zero rows, so the " +
                    "differential target compares nothing.");
            }
        }

        private static void VerifyOneEncryptedSeed(string seedPath)
        {
            string name = Path.GetFileName(seedPath);
            int rows;
            try
            {
                rows = Harnesses.OpenEncryptedSeedForSelfCheck(File.ReadAllBytes(seedPath));
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException(
                    $"the unmutated {name} no longer opens under Harnesses' encrypted options - the " +
                    "'encrypted' target would dead-end on this same rejection for every mutation too, " +
                    $"making it inert: {ex.GetType().FullName}: {ex.Message}",
                    ex);
            }
            if (rows == 0)
            {
                throw new InvalidOperationException(
                    $"the unmutated {name} opened but yielded zero rows under Harnesses' encrypted options.");
            }
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

        // Bytes that tend to sit right at a length/offset/count field's extremes — cheap to splice in
        // and far more likely to hit an unchecked boundary than a random byte run.
        private static readonly byte[][] _interesting =
        [
            [0x00], [0xFF], [0x7F], [0x80],
            [0xFF, 0xFF], [0xFF, 0xFF, 0xFF, 0xFF], [0x00, 0x00, 0x00, 0x80],
            [0xFF, 0xFF, 0xFF, 0x7F],
        ];

        // Blind byte-level mutations. Truncation matters most here: it is how a reader is made to
        // meet an offset or length field that points past the end of the data.
        private static byte[] Mutate(byte[] original, byte[][] corpus, Random random)
        {
            return random.Next(7) switch
            {
                0 => Truncate(original, random),
                1 => FlipBits(original, random),
                2 => FillRun(original, random),
                3 => SpliceWithin(original, random),
                4 => InsertRandom(original, random),
                5 => OverwriteWithBoundaryValue(original, random),
                _ => SpliceFromDonor(original, corpus, random),
            };
        }

        private static byte[] Truncate(byte[] original, Random random)
        {
            return original[..random.Next(0, original.Length + 1)];
        }

        private static byte[] FlipBits(byte[] original, Random random)
        {
            byte[] copy = [.. original];
            for (int i = 0; i < 8 && copy.Length > 0; i++)
            {
                int at = random.Next(copy.Length);
                copy[at] ^= (byte)(1 << random.Next(8));
            }
            return copy;
        }

        // Drives length/count fields to extremes.
        private static byte[] FillRun(byte[] original, Random random)
        {
            byte[] copy = [.. original];
            if (copy.Length > 0)
            {
                int start = random.Next(copy.Length);
                int length = Math.Min(copy.Length - start, random.Next(1, 17));
                byte value = (byte)random.Next(256);
                copy.AsSpan(start, length).Fill(value);
            }
            return copy;
        }

        private static byte[] SpliceWithin(byte[] original, Random random)
        {
            byte[] copy = [.. original];
            if (copy.Length > 4)
            {
                int length = random.Next(1, Math.Min(64, copy.Length));
                int from = random.Next(copy.Length - length + 1);
                int to = random.Next(copy.Length - length + 1);
                copy.AsSpan(from, length).CopyTo(copy.AsSpan(to));
            }
            return copy;
        }

        // The only mutation that grows the input; the rest shrink or hold size.
        private static byte[] InsertRandom(byte[] original, Random random)
        {
            int at = random.Next(original.Length + 1);
            int length = random.Next(1, 33);
            byte[] copy = new byte[original.Length + length];
            original.AsSpan(0, at).CopyTo(copy);
            random.NextBytes(copy.AsSpan(at, length));
            original.AsSpan(at).CopyTo(copy.AsSpan(at + length));
            return copy;
        }

        private static byte[] OverwriteWithBoundaryValue(byte[] original, Random random)
        {
            byte[] copy = [.. original];
            if (copy.Length > 0)
            {
                byte[] pattern = _interesting[random.Next(_interesting.Length)];
                int at = random.Next(copy.Length);
                int length = Math.Min(pattern.Length, copy.Length - at);
                pattern.AsSpan(0, length).CopyTo(copy.AsSpan(at, length));
            }
            return copy;
        }

        // Cross-format material: moves e.g. BIFF records from one container into another.
        private static byte[] SpliceFromDonor(byte[] original, byte[][] corpus, Random random)
        {
            byte[] donor = corpus[random.Next(corpus.Length)];
            byte[] copy = [.. original];
            if (copy.Length > 0 && donor.Length > 0)
            {
                int length = Math.Min(Math.Min(64, copy.Length), donor.Length);
                length = random.Next(1, length + 1);
                int from = random.Next(donor.Length - length + 1);
                int to = random.Next(copy.Length - length + 1);
                donor.AsSpan(from, length).CopyTo(copy.AsSpan(to));
            }
            return copy;
        }
    }
}
