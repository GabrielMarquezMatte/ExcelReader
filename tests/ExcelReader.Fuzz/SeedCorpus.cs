using System.Globalization;
using System.Text;
using ExcelReader.Core.Reader;
using ExcelReader.Core.Writer;

namespace ExcelReader.Fuzz
{
    /// <summary>
    /// Writes a starting corpus, one directory per fuzz target.
    /// </summary>
    /// <remarks>
    /// Generated rather than committed: a coverage-guided fuzzer needs a structurally valid starting
    /// point for each container (an .xlsb is a ZIP of BIFF12 parts — random bytes essentially never
    /// reach the record parsers on their own), and generating them keeps binaries out of the repo
    /// while guaranteeing the seeds always match the current writers. Seeds are kept deliberately
    /// tiny; libFuzzer explores small inputs far faster, and the real corpus grows from here.
    /// <para>
    /// Seeds are split per target because libFuzzer schedules and mutates every file in the corpus
    /// directory it is given. A flat corpus hands the csv target .xlsb blobs and the xlsx target
    /// plain text: those inputs die in the first format check, contribute no coverage, and still
    /// consume execution budget and corpus-scheduling weight for the whole run.
    /// </para>
    /// </remarks>
    internal static class SeedCorpus
    {
        private const string EncryptionPassword = "hunter2";

        private static readonly (string Target, string[] Seeds)[] _layout =
        [
            ("xlsx", ["xlsx", "xlsx-multisheet", "xlsx-sharedstrings", "xlsx-blanks", "xlsx-empty-sheet"]),
            ("xlsx-memory", ["xlsx", "xlsx-multisheet", "xlsx-sharedstrings", "xlsx-blanks", "xlsx-empty-sheet"]),
            ("xlsx-differential", ["xlsx", "xlsx-multisheet", "xlsx-sharedstrings", "xlsx-blanks", "xlsx-empty-sheet"]),
            ("xlsb", ["xlsb", "xlsb-multisheet"]),
            ("xlsb-memory", ["xlsb", "xlsb-multisheet"]),
            ("xlsb-differential", ["xlsb", "xlsb-multisheet"]),
            ("xls", ["xls"]),
            ("encrypted", ["encrypted", "encrypted-multisheet", "encrypted-sharedstrings"]),
            ("csv", ["csv", "csv-semicolon", "csv-tab", "csv-bom-lf", "csv-pipe", "csv-ragged",
                     "csv-cr-only", "csv-single-column", "csv-header-only", "csv-nul",
                     "csv-invalid-utf8", "csv-utf16le", "csv-utf16be", "csv-unterminated-quote"]),
            ("csv-sniff", ["csv", "csv-semicolon", "csv-tab", "csv-bom-lf", "csv-pipe", "csv-ragged",
                           "csv-cr-only", "csv-single-column", "csv-header-only", "csv-nul",
                           "csv-invalid-utf8", "csv-utf16le", "csv-utf16be", "csv-unterminated-quote"]),
            ("csv-parallel", ["csv", "csv-ragged", "csv-parallel-0", "csv-parallel-1", "csv-parallel-2",
                              "csv-parallel-3", "csv-parallel-4", "csv-parallel-5", "csv-parallel-6"]),
        ];

        internal static async Task GenerateAsync(string rootDirectory)
        {
            Dictionary<string, byte[]> seeds = await BuildAllAsync();

            int written = 0;
            foreach ((string target, string[] names) in _layout)
            {
                string directory = Path.Combine(rootDirectory, target);
                Directory.CreateDirectory(directory);
                foreach (string name in names)
                {
                    await File.WriteAllBytesAsync(Path.Combine(directory, $"seed-{name}.bin"), seeds[name]);
                    written++;
                }
            }
            Console.WriteLine(
                $"{written} seed(s) written across {_layout.Length} target director(ies) under {rootDirectory}");
        }

        private static async Task<Dictionary<string, byte[]>> BuildAllAsync()
        {
            byte[] xlsx = await XlsxAsync();
            byte[] xlsxMultiSheet = await MultiSheetAsync();
            byte[] xlsxSharedStrings = await SharedStringsAsync();

            var seeds = new Dictionary<string, byte[]>(StringComparer.Ordinal)
            {
                ["xlsx"] = xlsx,
                ["xlsx-multisheet"] = xlsxMultiSheet,
                ["xlsx-sharedstrings"] = xlsxSharedStrings,
                ["xlsx-blanks"] = await BlanksAsync(),
                ["xlsx-empty-sheet"] = await EmptySheetAsync(),
                ["xlsb"] = await XlsbAsync(),
                ["xlsb-multisheet"] = await MultiSheetXlsbAsync(),
                ["xls"] = await XlsAsync(),

                ["encrypted"] = Encrypt(xlsx),
                ["encrypted-multisheet"] = Encrypt(xlsxMultiSheet),
                ["encrypted-sharedstrings"] = Encrypt(xlsxSharedStrings),
            };

            foreach ((string name, byte[] bytes) in CsvSeeds())
            {
                seeds[name] = bytes;
            }
            return seeds;
        }

        private static byte[] Encrypt(byte[] package)
        {
            using var source = new MemoryStream(package, writable: false);
            using var destination = new MemoryStream();
            Excel.EncryptPackage(source, destination, ExcelPassword.FromString(EncryptionPassword));
            return destination.ToArray();
        }

        private static IEnumerable<(string Name, byte[] Bytes)> CsvSeeds()
        {
            return DialectSeeds().Concat(ParallelSeeds());
        }

        private static IEnumerable<(string Name, byte[] Bytes)> DialectSeeds()
        {
            yield return ("csv", Utf8(
                "name,qty,when\r\nplain,1,2024-01-02\r\n\"quo\"\"ted\",2,2024-01-03\r\n\"has,comma\",3,\r\n\"has\nnewline\",4,\r\n"));
            yield return ("csv-semicolon", Utf8(
                "name;qty;when\r\nplain;1;2024-01-02\r\n\"quo;ted\";2;2024-01-03\r\n"));
            yield return ("csv-tab", Utf8("name\tqty\twhen\nplain\t1\t2024-01-02\n"));
            yield return ("csv-pipe", Utf8("name|qty|when\nplain|1|2024-01-02\n\"a|b\"|2|\n"));

            yield return ("csv-bom-lf", Concat([0xEF, 0xBB, 0xBF], Utf8("name,qty\nplain,1\n\"quoted\",2\n")));

            yield return ("csv-cr-only", Utf8("name,qty\rplain,1\r\"quoted\",2\r"));

            yield return ("csv-ragged", Utf8("a,b,c\r\n1\r\n1,2\r\n1,2,3\r\n1,2,3,4,5,6\r\n,,\r\n"));

            yield return ("csv-single-column", Utf8("justonecolumn\nvalue\nanother\n"));

            yield return ("csv-header-only", Utf8("a,b,c"));

            yield return ("csv-nul", Concat(Utf8("a,b\n1,"), [0x00], Utf8("2\n\"\n")));

            yield return ("csv-invalid-utf8", Concat(Utf8("a,b\n"), [0xC3, 0x28, 0x80, 0xFF, 0xE2, 0x82], Utf8("\n")));

            yield return ("csv-utf16le", Concat([0xFF, 0xFE], Encoding.Unicode.GetBytes("a,b\r\n1,2\r\n")));
            yield return ("csv-utf16be", Concat([0xFE, 0xFF], Encoding.BigEndianUnicode.GetBytes("a,b\r\n1,2\r\n")));

            yield return ("csv-unterminated-quote", Utf8("a,b\n\"never closed,1\n2,3\n"));

        }

        private static IEnumerable<(string Name, byte[] Bytes)> ParallelSeeds()
        {
            string[] parallel =
            [
                "\"a\nb\",1,x",
                "\"a\"\"b\nc\",2,y",
                "\"\",3,",
                "\"\n\n\n\",4,z",
                "a,5,\"unterminated",
                "\"" + new string('x', 4096) + "\n" + new string('y', 4096) + "\",6,w",
                string.Concat(Enumerable.Repeat("\"q\",1,\"r\"\n", 64)),
            ];
            for (int i = 0; i < parallel.Length; i++)
            {
                yield return ($"csv-parallel-{i.ToString(CultureInfo.InvariantCulture)}", Utf8(parallel[i]));
            }
        }

        private static byte[] Utf8(string text)
        {
            return Encoding.UTF8.GetBytes(text);
        }

        private static byte[] Concat(params byte[][] parts)
        {
            byte[] result = new byte[parts.Sum(static p => p.Length)];
            int at = 0;
            foreach (byte[] part in parts)
            {
                part.CopyTo(result, at);
                at += part.Length;
            }
            return result;
        }

        private static async Task<byte[]> XlsxAsync()
        {
            using var ms = new MemoryStream();
            await using (XlsxWorkbookWriter wb = XlsxWorkbookWriter.Create(ms, leaveOpen: true))
            {
                XlsxSheetWriter sheet = wb.AddSheet("S");
                sheet.SetColumnWidth(0, 12);
                await sheet.StartAsync();
                await WriteSampleRowsAsync(sheet);
                await sheet.EndAsync();
                await wb.EndAsync();
            }
            return ms.ToArray();
        }

        private static async Task<byte[]> XlsbAsync()
        {
            using var ms = new MemoryStream();
            await using (XlsbWorkbookWriter wb = XlsbWorkbookWriter.Create(ms, leaveOpen: true))
            {
                XlsbSheetWriter sheet = wb.AddSheet("S");
                sheet.SetColumnWidth(0, 12);
                await sheet.StartAsync();
                await WriteSampleRowsAsync(sheet);
                await sheet.EndAsync();
                await wb.EndAsync();
            }
            return ms.ToArray();
        }

        private static async Task<byte[]> XlsAsync()
        {
            using var ms = new MemoryStream();
            await using (XlsWorkbookWriter wb = XlsWorkbookWriter.Create(ms, leaveOpen: true))
            {
                XlsSheetWriter sheet = wb.AddSheet("S");
                await sheet.StartAsync();
                await WriteSampleRowsAsync(sheet);
                await sheet.EndAsync();
                await wb.EndAsync();
            }
            return ms.ToArray();
        }

        private static async Task<byte[]> MultiSheetAsync()
        {
            using var ms = new MemoryStream();
            await using (XlsxWorkbookWriter wb = XlsxWorkbookWriter.Create(ms, leaveOpen: true))
            {
                foreach (string name in new[] { "First", "Second" })
                {
                    XlsxSheetWriter sheet = wb.AddSheet(name);
                    await sheet.StartAsync();
                    await WriteSampleRowsAsync(sheet);
                    await sheet.EndAsync();
                }
                await wb.EndAsync();
            }
            return ms.ToArray();
        }

        private static async Task<byte[]> MultiSheetXlsbAsync()
        {
            using var ms = new MemoryStream();
            await using (XlsbWorkbookWriter wb = XlsbWorkbookWriter.Create(ms, leaveOpen: true))
            {
                foreach (string name in new[] { "First", "Second" })
                {
                    XlsbSheetWriter sheet = wb.AddSheet(name);
                    await sheet.StartAsync();
                    await WriteSampleRowsAsync(sheet);
                    await sheet.EndAsync();
                }
                await wb.EndAsync();
            }
            return ms.ToArray();
        }

        private static async Task<byte[]> SharedStringsAsync()
        {
            using var ms = new MemoryStream();
            await using (XlsxWorkbookWriter wb = XlsxWorkbookWriter.Create(ms, leaveOpen: true))
            {
                XlsxSheetWriter sheet = wb.AddSheet("S");
                await sheet.StartAsync();
                for (int i = 0; i < 32; i++)
                {
                    await using XlsxRowWriter row = await sheet.StartRowAsync();
                    row.Write(i % 3 == 0 ? "repeated" : $"unique{i.ToString(CultureInfo.InvariantCulture)}");
                    row.Write(string.Empty);
                }
                await sheet.EndAsync();
                await wb.EndAsync();
            }
            return ms.ToArray();
        }

        private static async Task<byte[]> BlanksAsync()
        {
            using var ms = new MemoryStream();
            await using (XlsxWorkbookWriter wb = XlsxWorkbookWriter.Create(ms, leaveOpen: true))
            {
                XlsxSheetWriter sheet = wb.AddSheet("S");
                await sheet.StartAsync();
                await using (XlsxRowWriter row = await sheet.StartRowAsync())
                {
                    row.Write(value: (string?)null);
                    row.Write(value: (double?)null);
                    row.Write(value: (DateTime?)null);
                    row.Write(value: (bool?)null);
                }
                await sheet.EndAsync();
                await wb.EndAsync();
            }
            return ms.ToArray();
        }

        private static async Task<byte[]> EmptySheetAsync()
        {
            using var ms = new MemoryStream();
            await using (XlsxWorkbookWriter wb = XlsxWorkbookWriter.Create(ms, leaveOpen: true))
            {
                XlsxSheetWriter sheet = wb.AddSheet("S");
                await sheet.StartAsync();
                await sheet.EndAsync();
                await wb.EndAsync();
            }
            return ms.ToArray();
        }

        private static async Task WriteSampleRowsAsync<TRow>(ISheetWriter<TRow> sheet)
            where TRow : IRowWriter, IAsyncDisposable
        {
            await using (TRow header = await sheet.StartRowAsync())
            {
                header.Write("text");
                header.Write("number");
                header.Write("date");
                header.Write("bool");
            }
            await using (TRow row = await sheet.StartRowAsync())
            {
                row.Write("shared");
                row.Write(1234.5);
                row.Write(new DateTime(2024, 1, 2, 3, 4, 5, DateTimeKind.Unspecified));
                row.Write(value: true);
            }
        }
    }
}
