namespace ExcelReader.Tests.Crypto
{
    internal static class EncryptedFixtures
    {
        internal const string Password = "hunter2";

        internal static readonly string[] All =
        [
            "agile-aes256-sha512.xlsx",
            "agile-aes256-sha512.xlsb",
            "standard-aes128-sha1.xlsx",
        ];

        private static readonly Dictionary<string, string> Passwords = new(StringComparer.Ordinal)
        {
            ["standard-aes128-sha1.xlsx"] = "VelvetSweatshop",
        };

        internal static string PasswordFor(string name)
        {
            return Passwords.TryGetValue(name, out string? password) ? password : Password;
        }

        internal static string Dir => Path.Combine(AppContext.BaseDirectory, "data", "encrypted");

        internal static string Path_(string name)
        {
            return Path.Combine(Dir, name);
        }

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
