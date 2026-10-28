namespace Snippets.UntrustedInput.Options
{
    internal static class Snippet
    {
        internal static ExcelReaderOptions Run()
        {
            #region untrusted-input-options
            var uploadOptions = new ExcelReaderOptions
            {
                MaxTotalDecompressedBytes = 64L * 1024 * 1024, // everything one workbook may inflate
                MaxSharedStringBytes = 16L * 1024 * 1024,
                MaxCellBytes = 1024 * 1024,
                MaxZipEntries = 1_000,
            };
            #endregion
            return uploadOptions;
        }
    }
}

namespace Snippets.UntrustedInput.Reject
{
    internal static class Snippet
    {
        internal static void Run(Stream upload, ExcelReaderOptions uploadOptions)
        {
            #region untrusted-input-reject
            try
            {
                using IExcelWorkbook workbook = Excel.Open(upload, options: uploadOptions);
                foreach (Row row in workbook.FirstSheet)
                {
                    // stage the row; commit only once the loop has finished
                }
            }
            catch (ExcelLimitExceededException ex)
            {
                Console.WriteLine($"Rejected: {ex.LimitName} reached {ex.Actual}, limit {ex.Limit}");
            }
            catch (Exception ex) when (IsRejectedFile(ex))
            {
                Console.WriteLine($"Rejected: {ex.Message}");
            }

            static bool IsRejectedFile(Exception ex)
            {
                return ex is InvalidDataException or EndOfStreamException or NotSupportedException
                    or ExcelEncryptionException or ExcelParseException
                || ex.GetType() == typeof(ArgumentException)
                || ex.GetType() == typeof(InvalidOperationException);
            }
            #endregion
        }
    }
}

namespace Snippets.UntrustedInput.CsvInjection
{
    internal static class Snippet
    {
        internal static void Run(CsvRowWriter row, string comment)
        {
            #region untrusted-input-csv-injection
            static string Neutralize(string text)
            {
                return text.Length > 0 && "=+-@\t\r".Contains(text[0]) ? "'" + text : text;
            }

            row.Write(Neutralize(comment));
            #endregion
        }
    }
}
