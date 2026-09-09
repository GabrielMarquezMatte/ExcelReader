namespace ExcelReader.Tests
{
    // Encrypted fixtures are the only *third-party* oracle for decryption correctness -
    // Excel.EncryptPackage is round-tripped through the reader elsewhere in the test suite, but
    // that only proves encrypt and decrypt agree with each other, not independent correctness (see
    // data/encrypted/README.md). Each encrypted file has a paired ".plain." file produced by an
    // implementation this codebase did not write. See data/encrypted/README.md for what
    // schemes/key sizes this corpus does and does not cover.
    internal static class EncryptedFixtures
    {
        internal const string Password = "hunter2";

        internal static readonly string[] All =
        [
            "agile-aes256-sha512.xlsx",
            "agile-aes256-sha512.xlsb",
            "standard-aes128-sha1.xlsx",
        ];

        // Every fixture but one is encrypted with Password. "standard-aes128-sha1.xlsx" is a genuine
        // third-party file (Apache POI's test corpus) and keeps its real password instead.
        internal static string PasswordFor(string name)
        {
            return name.StartsWith("standard-", StringComparison.Ordinal) ? "VelvetSweatshop" : Password;
        }

        internal static string Dir => Path.Combine(AppContext.BaseDirectory, "data", "encrypted");

        internal static string Path_(string name)
        {
            return Path.Combine(Dir, name);
        }

        // "agile-aes256-sha512.xlsx" -> "agile-aes256-sha512.plain.xlsx"
        internal static string PlainPath(string name)
        {
            string ext = Path.GetExtension(name);
            return Path.Combine(Dir, Path.GetFileNameWithoutExtension(name) + ".plain" + ext);
        }

        internal static byte[] Bytes(string name)
        {
            return File.ReadAllBytes(Path_(name));
        }

        internal static byte[] PlainBytes(string name)
        {
            return File.ReadAllBytes(PlainPath(name));
        }
    }
}
