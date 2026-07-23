using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using ExcelReader.Core.Reader;
using ExcelReader.Core.Writer;

namespace ExcelReader.Tests
{
    // Seeded-random-mutator fuzz harness (F14 in docs/road-to-a.md). No binary-format parser in this
    // codebase (OLE/CFB, BIFF8, BIFF12, ZIP) had ever been exercised against randomized corruption —
    // only hand-crafted malformed inputs. This flips random bytes in otherwise-valid seed files and
    // requires every resulting failure to surface as one of this library's own graceful rejections
    // (ExcelLimitExceededException, or a well-known BCL parsing exception like InvalidDataException),
    // never an unhandled crash (NullReferenceException, IndexOutOfRangeException, etc.) that would
    // indicate a real bounds/validation bug reachable from untrusted input.
    //
    // The seed is fixed so a failure is 100% reproducible: re-running reproduces the exact same
    // mutations. If this ever fails, the byte offsets in the exception message plus MutateCopy's
    // logic are enough to reconstruct the corrupt file by hand.
    public class FuzzTests
    {
        private const int Seed = 20260723;
        private const int RoundsPerFormat = 500;

        // Anything NOT an instance of one of these types is treated as a real bug: a corrupt file must
        // be rejected gracefully, not crash the parser. Deliberately broad (covers the BCL exceptions
        // ZipArchive/Stream plumbing can itself throw on a corrupted container) since the goal is
        // catching genuine unhandled-crash bugs, not policing exact exception taxonomy.
        private static readonly Type[] AcceptableExceptionTypes =
        [
            typeof(InvalidDataException),
            typeof(ExcelLimitExceededException),
            typeof(EndOfStreamException),
            typeof(IOException),
            typeof(OverflowException),
            typeof(ArgumentException), // covers ArgumentOutOfRangeException, ArgumentNullException
            typeof(NotSupportedException),
            typeof(FormatException),
        ];

        [Fact]
        public void MutatedXlsxBytesNeverCrashTheReader()
        {
            byte[] seed = BuildXlsxSeed();
            int completed = FuzzFormat(seed, format: "xlsx");
            Assert.Equal(RoundsPerFormat, completed);
        }

        [Fact]
        public async Task MutatedXlsbBytesNeverCrashTheReader()
        {
            byte[] seed = await BuildXlsbSeedAsync();
            int completed = FuzzFormat(seed, format: "xlsb");
            Assert.Equal(RoundsPerFormat, completed);
        }

        [Fact]
        public void MutatedXlsBytesNeverCrashTheReader()
        {
            byte[] seed = BuildXlsSeed();
            int completed = FuzzFormat(seed, format: "xls");
            Assert.Equal(RoundsPerFormat, completed);
        }

        // Returns the number of rounds completed (always RoundsPerFormat unless it throws first).
        // Any exception escaping a round that isn't in AcceptableExceptionTypes is rewrapped with the
        // mutated byte offsets so the failure is reproducible, then rethrown — that failure is what
        // fails the calling [Fact].
        private static int FuzzFormat(byte[] seed, string format)
        {
            var rng = new Random(Seed);
            for (int round = 0; round < RoundsPerFormat; round++)
            {
                byte[] mutated = MutateCopy(seed, rng, out int[] positions);
                try
                {
                    OpenAndDrain(mutated);
                }
                catch (Exception ex) when (IsAcceptable(ex))
                {
                    // Expected: the mutated bytes were rejected gracefully.
                }
                catch (Exception ex)
                {
                    string offsets = string.Join(", ", positions);
                    throw new InvalidOperationException(
                        string.Create(CultureInfo.InvariantCulture,
                            $"Round {round} on {format} seed produced an unhandled '{ex.GetType().Name}' (mutated byte offsets: [{offsets}]). This indicates a validation/bounds gap reachable from untrusted input, not a graceful rejection."),
                        ex);
                }
            }
            return RoundsPerFormat;
        }

        private static void OpenAndDrain(byte[] bytes)
        {
            using var ms = new MemoryStream(bytes);
            using IExcelRowReader reader = Excel.Open(ms);
            for (int s = 0; s < reader.SheetCount; s++)
            {
                reader.MoveToSheet(s);
                using IExcelRowEnumerator e = reader.GetEnumerator();
                while (e.MoveNext())
                {
                    for (int c = 0; c < e.Current.ColumnCount; c++)
                    {
                        _ = e.Current[c].GetString();
                    }
                }
            }
        }

        private static bool IsAcceptable(Exception ex)
        {
            foreach (Type acceptableType in AcceptableExceptionTypes)
            {
                if (acceptableType.IsInstanceOfType(ex))
                {
                    return true;
                }
            }
            return false;
        }

        // Flips 1-8 random bytes to random values; a copy so the seed itself is never mutated.
        // Deliberately not cryptographically secure — CA5394 doesn't apply: a seeded, reproducible
        // PRNG is exactly what makes a fuzz failure pinpoint-able, unlike a CSPRNG would be.
        [SuppressMessage("Security", "CA5394:Do not use insecure randomness",
            Justification = "Fuzzing needs a reproducible seeded PRNG, not cryptographic randomness.")]
        [SuppressMessage("Performance", "HLQ013:Consider using 'foreach' loop instead of 'for' loop",
            Justification = "Each iteration both reads (rng.Next) and writes positions[i] by index; foreach can't express the write.")]
        private static byte[] MutateCopy(byte[] seed, Random rng, out int[] positions)
        {
            byte[] copy = (byte[])seed.Clone();
            int count = rng.Next(1, 9);
            positions = new int[count];
            for (int i = 0; i < count; i++)
            {
                int pos = rng.Next(copy.Length);
                positions[i] = pos;
                copy[pos] = (byte)rng.Next(256);
            }
            return copy;
        }

        private static byte[] BuildXlsxSeed()
        {
            using MemoryStream ms = WorkbookBuilder.BuildMultiSheet(
                sheets:
                [
                    ("S1", """<row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1"><v>42</v></c></row><row r="2"><c r="A2" t="s"><v>1</v></c></row>"""),
                    ("S2", """<row r="1"><c r="A1"><v>7</v></c></row>"""),
                ],
                sharedStrings: "<si><t>hello</t></si><si><t>world</t></si>",
                styles: "<styleSheet><cellXfs count=\"1\"><xf numFmtId=\"14\"/></cellXfs></styleSheet>");
            return ms.ToArray();
        }

        private static async Task<byte[]> BuildXlsbSeedAsync()
        {
            CancellationToken ct = TestContext.Current.CancellationToken;
            MemoryStream ms = new();
            await using (XlsbWorkbookWriter wb = await XlsbWorkbookWriter.CreateAsync(ms, leaveOpen: true, ct: ct))
            {
                await wb.StartAsync(ct);
                XlsbSheetWriter sheet = wb.AddSheet("Sheet1");
                await sheet.StartAsync(ct);
                await using (XlsbRowWriter row = await sheet.StartRowAsync(ct))
                {
                    row.Write("hello");
                    row.Write(42);
                    row.Write(true);
                    row.Write(new DateTime(2026, 1, 1));
                }
                await sheet.EndAsync(ct);
                await wb.EndAsync(ct);
            }
            return ms.ToArray();
        }

        private static byte[] BuildXlsSeed()
        {
            using MemoryStream ms = XlsWorkbookBuilder.Build(sheets:
            [
                ("S1", [["Name", 1, true], ["Alice", 2, false]]),
            ]);
            return ms.ToArray();
        }
    }
}
