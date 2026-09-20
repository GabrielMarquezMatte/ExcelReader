using ExcelReader.Core.Reader;

namespace ExcelReader.Core.Crypto
{
    internal abstract class PackageCipher : IDisposable
    {
        internal static PackageCipher Create(EncryptionDescriptor descriptor, ExcelPassword password)
        {
            return descriptor switch
            {
                AgileDescriptor agile => new AgilePackageCipher(agile, password.Chars),
                StandardDescriptor standard => new StandardPackageCipher(standard, password.Chars),
                _ => throw new ExcelEncryptionException(ExcelEncryptionReason.UnsupportedScheme,
                    "This workbook's encryption scheme has no decryption implementation."),
            };
        }

        internal abstract bool SupportsIntegrity { get; }

        internal abstract void DecryptSegment(int segmentIndex, ReadOnlyMemory<byte> cipher, Memory<byte> plain);

        internal abstract void VerifyIntegrity(Stream ciphertextView);

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected abstract void Dispose(bool disposing);
    }
}
