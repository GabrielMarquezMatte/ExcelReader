namespace ExcelReader.Tests
{
    // Encrypted fixtures are the only oracle for decryption correctness (writing encrypted files is
    // out of scope, so there is no round-trip check). Each encrypted file has a paired ".plain."
    // file produced by msoffcrypto-tool, an independent implementation. See data/encrypted/README.md
    // for what schemes/key sizes this corpus does and does not cover.
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
