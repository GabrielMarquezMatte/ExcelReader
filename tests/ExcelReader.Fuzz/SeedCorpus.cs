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
