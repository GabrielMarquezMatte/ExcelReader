using System.Globalization;
using System.IO.Compression;
using ExcelReader.Core.Reader;
using ExcelReader.Core.Writer.Xlsb;
using ExcelReader.Tests.Reader.Xls;

namespace ExcelReader.Tests
{
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

        [Fact]
        public void MutatedXlsBytesWithContinueRecordNeverCrashTheReader()
        {
            byte[] seed = BuildXlsSeedWithContinuedSst();
            int completed = FuzzFormat(seed, format: "xls-continue");
            Assert.Equal(RoundsPerFormat, completed);
        }

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
            await using (XlsbWorkbookWriter wb = XlsbWorkbookWriter.Create(ms, leaveOpen: true))
            {
                XlsbSheetWriter sheet = wb.AddSheet("Sheet1");
                await using (XlsbRowWriter row = await sheet.StartRowAsync(ct))
                {
                    row.Write("hello");
                    row.Write(new string('x', 1000));
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
            byte[] firstRegion =
            [
                0x02, 0x00, 0x00, (byte)'A', (byte)'B',
                0x04, 0x00, 0x00, (byte)'C', (byte)'D',
            ];
            byte[] continueRegion = [0x00, (byte)'E', (byte)'F'];
            byte[] framed = XlsWorkbookBuilder.FrameSstWithContinue(cstTotal: 2, cstUnique: 2, firstRegion, continueRegion);
            using MemoryStream ms = XlsWorkbookBuilder.BuildRawSst(framed, labelSstCount: 2);
            return ms.ToArray();
        }
    }
}
