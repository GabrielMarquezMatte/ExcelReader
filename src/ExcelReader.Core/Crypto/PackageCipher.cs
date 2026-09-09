using ExcelReader.Core.Reader;

namespace ExcelReader.Core.Crypto
{
    // One keyed decryption strategy per encryption scheme. The package readers hold a cipher rather
    // than a descriptor so they never name a scheme: they hand it one aligned ciphertext window and
    // get plaintext back. Every scheme's segment size is the caller's 4096-byte choice, which works
    // because agile derives its IV from the segment index and standard is ECB, where blocks are
    // independent.
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

        // False for a scheme with no integrity field, and for an agile file whose writer omitted
        // <dataIntegrity>.
        internal abstract bool SupportsIntegrity { get; }

        // `cipher` and `plain` are equal-length and always a whole number of cipher blocks: the
        // caller validates that the total ciphertext length is block-aligned, and the segment size
        // is itself a multiple of the block size, so even a short final segment stays aligned.
        internal abstract void DecryptSegment(int segmentIndex, ReadOnlySpan<byte> cipher, Span<byte> plain);

        // Costs a full pass over the ciphertext, so the caller decides whether to pay for it.
        internal abstract void VerifyIntegrity(Stream ciphertextView);

        // Standard Dispose(bool) shape: matches this codebase's other abstract-base-with-subclasses
        // IDisposable types (e.g. NativeWriterHandle), which the Sonar/CA dispose-pattern analyzers
        // require of an inheritable class holding an IDisposable implementation directly.
        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected abstract void Dispose(bool disposing);
    }
}
