namespace Snippets.Encryption.Open
{
    internal static class Snippet
    {
        internal static void Run()
        {
            #region encryption-open
            var options = new ExcelReaderOptions { Password = "hunter2" };
            using IExcelWorkbook workbook = Excel.Open("protected.xlsx", options);
            foreach (Row row in workbook.FirstSheet) { /* ... */ }
            #endregion
        }
    }
}

namespace Snippets.Encryption.Reason
{
    internal static class Snippet
    {
        internal static void Run(string path, ExcelReaderOptions options)
        {
            #region encryption-reason
            try
            {
                using IExcelWorkbook workbook = Excel.Open(path, options);
            }
            catch (ExcelEncryptionException ex) when (ex.Reason is ExcelEncryptionReason.PasswordIncorrect)
            {
                // Ask again. UnsupportedScheme and IntegrityFailure are terminal.
            }
            #endregion
        }
    }
}

namespace Snippets.Encryption.VerifyIntegrity
{
    internal static class Snippet
    {
        internal static void Run()
        {
            #region encryption-verify-integrity
            var options = new ExcelReaderOptions
            {
                Password = "hunter2",
                VerifyEncryptedIntegrity = true, // pay one full pass; reject tampered packages
            };
            #endregion
        }
    }
}

namespace Snippets.Encryption.Write
{
    #region encryption-write
    using ExcelReader.Core.Reader;
    using ExcelReader.Core.Writer;
    using ExcelReader.Core.Writer.Xlsx;
    #endregion

    internal static class Snippet
    {
        internal static async Task Run()
        {
            #region encryption-write
            using var plain = new MemoryStream();
            await using (var workbook = XlsxWorkbookWriter.Create(plain, leaveOpen: true))
            {
                // write sheets and rows as usual
            }
            plain.Position = 0;

            using var destination = File.Create("secret.xlsx");
            await Excel.EncryptPackageAsync(plain, destination, "hunter2");
            #endregion
        }
    }
}
