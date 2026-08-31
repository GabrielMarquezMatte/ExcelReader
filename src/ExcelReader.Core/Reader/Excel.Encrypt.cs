using ExcelReader.Core.Crypto;

namespace ExcelReader.Core.Reader
{
    public static partial class Excel
    {
        /// <summary>
        /// Wraps a finished plaintext OOXML package (XLSX or XLSB) in a password-encrypted CFB
        /// container, using agile encryption (ECMA-376 4.4): AES-256-CBC, SHA-512, 100,000 spin
        /// iterations, with a <c>dataIntegrity</c> HMAC.
        /// </summary>
        /// <param name="package">
        /// The plaintext package, positioned at its first byte. Must be readable and seekable — it is
        /// read twice. Not disposed by this method.
        /// </param>
        /// <param name="destination">Where the encrypted container is written. Need not be seekable. Not disposed by this method.</param>
        /// <param name="password">The password required to open the result.</param>
        /// <exception cref="ArgumentNullException">Any argument is <see langword="null"/>.</exception>
        /// <exception cref="ArgumentException">
        /// <paramref name="package"/> is unreadable, not seekable, empty, or does not begin with the
        /// <c>PK\x03\x04</c> package signature; or <paramref name="password"/> is empty.
        /// </exception>
        public static void EncryptPackage(Stream package, Stream destination, ExcelPassword password)
        {
            ArgumentNullException.ThrowIfNull(destination);
            PackageEncryptor.Encrypt(package, destination, password);
        }

        /// <summary>
        /// Reads the plaintext package at <paramref name="packagePath"/> and writes its encrypted
        /// counterpart to <paramref name="destinationPath"/>, overwriting an existing file.
        /// </summary>
        /// <param name="packagePath">Path to the plaintext XLSX/XLSB package.</param>
        /// <param name="destinationPath">Path to write the encrypted workbook to.</param>
        /// <param name="password">The password required to open the result.</param>
        /// <exception cref="ArgumentException">
        /// <paramref name="packagePath"/> or <paramref name="destinationPath"/> is <see langword="null"/> or empty;
        /// or the package at <paramref name="packagePath"/> is unreadable, empty, or does not begin with the
        /// <c>PK\x03\x04</c> package signature; or <paramref name="password"/> is empty.
        /// </exception>
        public static void EncryptPackage(string packagePath, string destinationPath, ExcelPassword password)
        {
            ArgumentException.ThrowIfNullOrEmpty(packagePath);
            ArgumentException.ThrowIfNullOrEmpty(destinationPath);
            using FileStream package = File.OpenRead(packagePath);
            using FileStream destination = new(destinationPath, FileMode.Create, FileAccess.Write, FileShare.None);
            PackageEncryptor.Encrypt(package, destination, password);
        }

        /// <summary>Asynchronous counterpart to <see cref="EncryptPackage(Stream, Stream, ExcelPassword)"/>.</summary>
        /// <param name="package">The plaintext package. Must be readable and seekable; not disposed here.</param>
        /// <param name="destination">Where the encrypted container is written; not disposed here.</param>
        /// <param name="password">The password required to open the result.</param>
        /// <param name="ct">A token to cancel the operation.</param>
        /// <exception cref="ArgumentNullException">Any argument is <see langword="null"/>.</exception>
        /// <exception cref="ArgumentException">
        /// <paramref name="package"/> is unreadable, not seekable, empty, or does not begin with the
        /// <c>PK\x03\x04</c> package signature; or <paramref name="password"/> is empty.
        /// </exception>
        public static ValueTask EncryptPackageAsync(Stream package, Stream destination, ExcelPassword password,
            CancellationToken ct = default)
        {
            ArgumentNullException.ThrowIfNull(destination);
            return PackageEncryptor.EncryptAsync(package, destination, password, ct);
        }

        /// <summary>Asynchronous counterpart to <see cref="EncryptPackage(string, string, ExcelPassword)"/>.</summary>
        /// <param name="packagePath">Path to the plaintext XLSX/XLSB package.</param>
        /// <param name="destinationPath">Path to write the encrypted workbook to.</param>
        /// <param name="password">The password required to open the result.</param>
        /// <param name="ct">A token to cancel the operation.</param>
        /// <exception cref="ArgumentException">
        /// <paramref name="packagePath"/> or <paramref name="destinationPath"/> is <see langword="null"/> or empty;
        /// or the package at <paramref name="packagePath"/> is unreadable, empty, or does not begin with the
        /// <c>PK\x03\x04</c> package signature; or <paramref name="password"/> is empty.
        /// </exception>
        public static async ValueTask EncryptPackageAsync(string packagePath, string destinationPath,
            ExcelPassword password, CancellationToken ct = default)
        {
            ArgumentException.ThrowIfNullOrEmpty(packagePath);
            ArgumentException.ThrowIfNullOrEmpty(destinationPath);
            FileStream package = new(packagePath, FileMode.Open, FileAccess.Read, FileShare.Read,
                bufferSize: 4096, useAsync: true);
            await using (package.ConfigureAwait(false))
            {
                FileStream destination = new(destinationPath, FileMode.Create, FileAccess.Write,
                    FileShare.None, bufferSize: 4096, useAsync: true);
                await using (destination.ConfigureAwait(false))
                {
                    await PackageEncryptor.EncryptAsync(package, destination, password, ct).ConfigureAwait(false);
                }
            }
        }
    }
}
