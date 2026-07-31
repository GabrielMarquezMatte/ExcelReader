using System.Globalization;
using System.IO.Compression;
using ExcelReader.Core.Reader;
using ExcelReader.Core.Writer;

namespace ExcelReader.Tests
{
    // Seeded-random-mutator fuzz harness. No binary-format parser in this codebase (OLE/CFB, BIFF8,
    // BIFF12, ZIP) had ever been exercised against randomized corruption —
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

        // The plain xls seed above never has its SST split across a CONTINUE record, so mutation never
        // touches that decode path (SEC-5: a zero-length CONTINUE can grow an unbounded list invisible
        // to the byte limit). This seed forces a real CONTINUE boundary using the same byte layout as
        // XlsReaderTests.SharedStringSplitAcrossContinueBoundaryDecodesCorrectly (a known-good, already
        // passing construction), so mutation now has that boundary to corrupt.
        [Fact]
        public void MutatedXlsBytesWithContinueRecordNeverCrashTheReader()
        {
            byte[] seed = BuildXlsSeedWithContinuedSst();
            int completed = FuzzFormat(seed, format: "xls-continue");
            Assert.Equal(RoundsPerFormat, completed);
        }

        // SEC-7 item 4: whole-file byte flips mostly corrupt the ZIP container itself (a bad CRC/local
        // header), so the mutation dies in DeflateStream/ZipArchive before it ever reaches this
        // library's own XML/BIFF12 parsing — the part these formats actually need fuzzed. Mutating one
        // part's decompressed content, then rebuilding a structurally valid ZIP with a fresh CRC via
        // ZipArchive, guarantees every round's corruption lands where the parser can see it.
        [Fact]
        public void MutatedXlsxSharedStringsContentNeverCrashesTheReader()
        {
            byte[] seed = BuildXlsxSeed();
            int completed = FuzzZipEntryContent(seed, "xl/sharedStrings.xml", format: "xlsx-sharedStrings");
            Assert.Equal(RoundsPerFormat, completed);
        }

        [Fact]
        public async Task MutatedXlsbWorkbookBinContentNeverCrashesTheReader()
        {
            byte[] seed = await BuildXlsbSeedAsync();
            int completed = FuzzZipEntryContent(seed, "xl/worksheets/sheet1.bin", format: "xlsb-sheet1");
            Assert.Equal(RoundsPerFormat, completed);
        }

        private static int FuzzZipEntryContent(byte[] seed, string entryName, string format)
        {
            byte[] entryContent = ReadZipEntry(seed, entryName);
            var rng = new Random(Seed);
            for (int round = 0; round < RoundsPerFormat; round++)
            {
                byte[] mutatedEntry = FuzzMutation.MutateCopy(entryContent, rng, out int[] positions);
                byte[] mutatedZip = RebuildZipWithEntry(seed, entryName, mutatedEntry);
                try
                {
                    FuzzMutation.RunBounded(() => OpenAndDrain(mutatedZip));
                }
                catch (Exception ex) when (FuzzMutation.IsAcceptable(ex))
                {
                    // Expected: the mutated bytes were rejected gracefully.
                }
                catch (Exception ex)
                {
                    string offsets = string.Join(", ", positions);
                    throw new InvalidOperationException(
                        string.Create(CultureInfo.InvariantCulture,
                            $"Round {round} on {format} seed (entry '{entryName}') produced an unhandled '{ex.GetType().Name}' (mutated byte offsets within the entry: [{offsets}]). This indicates a validation/bounds gap reachable from untrusted input, not a graceful rejection."),
                        ex);
                }
            }
            return RoundsPerFormat;
        }

        private static byte[] ReadZipEntry(byte[] zipBytes, string entryName)
        {
            using var input = new MemoryStream(zipBytes);
            using var zip = new ZipArchive(input, ZipArchiveMode.Read);
            using Stream entryStream = zip.GetEntry(entryName)!.Open();
            using var buffer = new MemoryStream();
            entryStream.CopyTo(buffer);
            return buffer.ToArray();
        }

        // Copies every entry verbatim except `entryName`, which is written with `newContent` instead —
        // ZipArchive computes a fresh CRC32/sizes for it, so the result is always a structurally valid
        // ZIP even though the part's content is corrupted.
        private static byte[] RebuildZipWithEntry(byte[] zipBytes, string entryName, byte[] newContent)
        {
            using var input = new MemoryStream(zipBytes);
            using var output = new MemoryStream();
            using (var srcZip = new ZipArchive(input, ZipArchiveMode.Read))
            using (var dstZip = new ZipArchive(output, ZipArchiveMode.Create, leaveOpen: true))
            {
                foreach (ZipArchiveEntry srcEntry in srcZip.Entries)
                {
                    ZipArchiveEntry dstEntry = dstZip.CreateEntry(srcEntry.FullName, CompressionLevel.Fastest);
                    using Stream dstStream = dstEntry.Open();
                    if (string.Equals(srcEntry.FullName, entryName, StringComparison.Ordinal))
                    {
                        dstStream.Write(newContent);
                    }
                    else
                    {
                        using Stream srcStream = srcEntry.Open();
                        srcStream.CopyTo(dstStream);
                    }
                }
            }
            return output.ToArray();
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
                byte[] mutated = FuzzMutation.MutateCopy(seed, rng, out int[] positions);
                try
                {
                    FuzzMutation.RunBounded(() => OpenAndDrain(mutated));
                }
                catch (Exception ex) when (FuzzMutation.IsAcceptable(ex))
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
                    row.Write(new string('x', 1000)); // long string
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

        private static byte[] BuildXlsSeedWithContinuedSst()
        {
            // string 0 = "AB" (cch=2, compressed); string 1 = "CDEF" split after "CD" — the SST record
            // ends mid-way through string 1's character array and a CONTINUE record resumes it.
            byte[] firstRegion =
            [
                0x02, 0x00, 0x00, (byte)'A', (byte)'B',
                0x04, 0x00, 0x00, (byte)'C', (byte)'D',
            ];
            byte[] continueRegion = [0x00, (byte)'E', (byte)'F']; // grbit + remaining two chars
            byte[] framed = XlsWorkbookBuilder.FrameSstWithContinue(cstTotal: 2, cstUnique: 2, firstRegion, continueRegion);
            using MemoryStream ms = XlsWorkbookBuilder.BuildRawSst(framed, labelSstCount: 2);
            return ms.ToArray();
        }
    }
}
