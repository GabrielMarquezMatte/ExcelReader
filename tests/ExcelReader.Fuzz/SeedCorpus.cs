using System.Globalization;
using ExcelReader.Core.Writer;

namespace ExcelReader.Fuzz
{
    /// <summary>
    /// Writes one small, valid file per format to use as fuzzing seeds.
    /// </summary>
    /// <remarks>
    /// Generated rather than committed: a coverage-guided fuzzer needs a structurally valid starting
    /// point for each container (an .xlsb is a ZIP of BIFF12 parts — random bytes essentially never
    /// reach the record parsers on their own), and generating them keeps binaries out of the repo
    /// while guaranteeing the seeds always match the current writers. Seeds are kept deliberately
    /// tiny; libFuzzer explores small inputs far faster, and the real corpus grows from here.
    /// </remarks>
    internal static class SeedCorpus
    {
        internal static async Task GenerateAsync(string directory)
        {
            Directory.CreateDirectory(directory);
            await File.WriteAllBytesAsync(Path.Combine(directory, "seed.xlsx"), await XlsxAsync());
            await File.WriteAllBytesAsync(Path.Combine(directory, "seed.xlsb"), await XlsbAsync());
            await File.WriteAllBytesAsync(Path.Combine(directory, "seed.xls"), await XlsAsync());
            await File.WriteAllBytesAsync(Path.Combine(directory, "seed.csv"), Csv());

            // Extra seed shapes: a coverage-guided fuzzer only ever explores paths a seed already
            // touches, so each of these exercises one shape the "happy path" seeds above never do
            // (multiple sheets, a shared string repeated enough to build a real dictionary, null/blank
            // cells, an empty sheet, and CSV dialects other than comma/CRLF).
            await File.WriteAllBytesAsync(Path.Combine(directory, "seed-multisheet.xlsx"), await MultiSheetAsync());
            await File.WriteAllBytesAsync(Path.Combine(directory, "seed-multisheet.xlsb"), await MultiSheetXlsbAsync());
            await File.WriteAllBytesAsync(Path.Combine(directory, "seed-sharedstrings.xlsx"), await SharedStringsAsync());
            await File.WriteAllBytesAsync(Path.Combine(directory, "seed-blanks.xlsx"), await BlanksAsync());
            await File.WriteAllBytesAsync(Path.Combine(directory, "seed-empty-sheet.xlsx"), await EmptySheetAsync());
            await File.WriteAllBytesAsync(Path.Combine(directory, "seed-semicolon.csv"), CsvSemicolon());
            await File.WriteAllBytesAsync(Path.Combine(directory, "seed-tab.csv"), CsvTab());
            await File.WriteAllBytesAsync(Path.Combine(directory, "seed-bom-lf.csv"), CsvBomLf());
            Console.WriteLine($"seeds written to {directory}");
        }

        private static async Task<byte[]> XlsxAsync()
        {
            using var ms = new MemoryStream();
            await using (XlsxWorkbookWriter wb = await XlsxWorkbookWriter.CreateAsync(ms, leaveOpen: true))
            {
                await wb.StartAsync();
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
            await using (XlsbWorkbookWriter wb = await XlsbWorkbookWriter.CreateAsync(ms, leaveOpen: true))
            {
                await wb.StartAsync();
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
                await wb.StartAsync();
                XlsSheetWriter sheet = wb.AddSheet("S");
                await sheet.StartAsync();
                await WriteSampleRowsAsync(sheet);
                await sheet.EndAsync();
                await wb.EndAsync();
            }
            return ms.ToArray();
        }

        private static byte[] Csv()
        {
            // Covers the shapes the CSV parser branches on: plain, quoted, embedded quote/delimiter,
            // embedded newline, and a CRLF terminator.
            const string text = "name,qty,when\r\nplain,1,2024-01-02\r\n\"quo\"\"ted\",2,2024-01-03\r\n\"has,comma\",3,\r\n\"has\nnewline\",4,\r\n";
            return System.Text.Encoding.UTF8.GetBytes(text);
        }

        private static byte[] CsvSemicolon()
        {
            const string text = "name;qty;when\r\nplain;1;2024-01-02\r\n\"quo;ted\";2;2024-01-03\r\n";
            return System.Text.Encoding.UTF8.GetBytes(text);
        }

        private static byte[] CsvTab()
        {
            const string text = "name\tqty\twhen\nplain\t1\t2024-01-02\n";
            return System.Text.Encoding.UTF8.GetBytes(text);
        }

        // UTF-8 BOM plus LF-only line endings: two dialect axes CsvSniffer branches on that the
        // primary seed (CRLF, no BOM) never exercises.
        private static byte[] CsvBomLf()
        {
            const string text = "name,qty\nplain,1\n\"quoted\",2\n";
            byte[] bom = [0xEF, 0xBB, 0xBF];
            byte[] body = System.Text.Encoding.UTF8.GetBytes(text);
            byte[] result = new byte[bom.Length + body.Length];
            bom.CopyTo(result, 0);
            body.CopyTo(result, bom.Length);
            return result;
        }

        // Two sheets: sheet-index bookkeeping (offsets/pointers into a sheet directory) is only
        // exercised once there is more than one sheet to point past.
        private static async Task<byte[]> MultiSheetAsync()
        {
            using var ms = new MemoryStream();
            await using (XlsxWorkbookWriter wb = await XlsxWorkbookWriter.CreateAsync(ms, leaveOpen: true))
            {
                await wb.StartAsync();
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
            await using (XlsbWorkbookWriter wb = await XlsbWorkbookWriter.CreateAsync(ms, leaveOpen: true))
            {
                await wb.StartAsync();
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

        // The same string repeated across many rows: forces a real shared-string dictionary with
        // duplicate entries, instead of the one-string-per-cell table the plain seed builds.
        private static async Task<byte[]> SharedStringsAsync()
        {
            using var ms = new MemoryStream();
            await using (XlsxWorkbookWriter wb = await XlsxWorkbookWriter.CreateAsync(ms, leaveOpen: true))
            {
                await wb.StartAsync();
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

        // Null cells interleaved with values: null-handling branches (blank vs. missing vs. typed)
        // never fire if every cell in the seed corpus is populated.
        private static async Task<byte[]> BlanksAsync()
        {
            using var ms = new MemoryStream();
            await using (XlsxWorkbookWriter wb = await XlsxWorkbookWriter.CreateAsync(ms, leaveOpen: true))
            {
                await wb.StartAsync();
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

        // A workbook whose only sheet has zero rows: SheetCount/MoveToSheet bookkeeping should still
        // hold with nothing to enumerate.
        private static async Task<byte[]> EmptySheetAsync()
        {
            using var ms = new MemoryStream();
            await using (XlsxWorkbookWriter wb = await XlsxWorkbookWriter.CreateAsync(ms, leaveOpen: true))
            {
                await wb.StartAsync();
                XlsxSheetWriter sheet = wb.AddSheet("S");
                await sheet.StartAsync();
                await sheet.EndAsync();
                await wb.EndAsync();
            }
            return ms.ToArray();
        }

        // A header plus one row of every cell kind the readers decode differently.
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
