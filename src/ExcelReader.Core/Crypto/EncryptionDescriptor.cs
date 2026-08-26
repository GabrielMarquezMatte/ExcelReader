using System.Buffers.Binary;
using System.Globalization;
using System.Xml;
using ExcelReader.Core.Reader;

namespace ExcelReader.Core.Crypto
{
    // Parses the CFB "EncryptionInfo" stream: an 8-byte little-endian (major, minor) version tuple,
    // followed by either a UTF-8 XML descriptor (agile, 4.4) or a binary header+verifier (standard,
    // 3.2/4.2). This runs on wholly untrusted input, BEFORE any password check, so every value read
    // is bounded before use and dispatched through an explicit allowlist rather than a reflected
    // lookup (see ParseHash below).
    //
    // ParseAgile uses XmlReader — the one deliberate departure from this codebase's hand-rolled XML
    // custom (XlsxXml, TagSpanEnumerable). It's justified here: the descriptor is roughly 1 KB, parsed
    // exactly once on a cold path, and a mis-read base64 attribute would fail silently as a wrong key
    // rather than a visible parse error, unlike the hot-path worksheet scanners this codebase otherwise
    // hand-rolls for throughput. XmlReader is BCL and AOT-safe. Because the XML is untrusted and parsed
    // pre-authentication, DTD processing and external entity resolution are hard-disabled — see
    // XmlSettings.
    internal abstract record EncryptionDescriptor
    {
        // [MS-OFFCRYPTO] fixes these namespace URIs; they identify the descriptor's default elements
        // and the password key encryptor regardless of whatever prefix a producer chooses for them.
        private const string EncryptionNamespace = "http://schemas.microsoft.com/office/2006/encryption";
        private const string PasswordKeyEncryptorNamespace = "http://schemas.microsoft.com/office/2006/keyEncryptor/password";

        // This descriptor is parsed BEFORE any password check, on wholly untrusted input, so DTD
        // processing and external resolution must be off - an XXE here would be reachable
        // pre-authentication.
        private static readonly XmlReaderSettings XmlSettings = new()
        {
            DtdProcessing = DtdProcessing.Prohibit,
            XmlResolver = null,
            IgnoreComments = true,
            IgnoreProcessingInstructions = true,
            IgnoreWhitespace = true,
            CloseInput = true,
        };

        internal static EncryptionDescriptor Parse(ReadOnlySpan<byte> info, ExcelReaderOptions options)
        {
            if (info.Length < 8)
            {
                throw new InvalidDataException("The EncryptionInfo stream is truncated.");
            }
            int major = BinaryPrimitives.ReadUInt16LittleEndian(info);
            int minor = BinaryPrimitives.ReadUInt16LittleEndian(info[2..]);
            return (major, minor) switch
            {
                (4, 4) => ParseAgile(info[8..], options),
                // Recognized but not implemented in this pass (see the plan's "Execution Scope Note" /
                // Task 17): no real standard-encryption fixture exists yet to verify a derivation
                // against, so this is a clean, honest rejection rather than an unverified crypto path.
                (4, 2) or (3, 2) => throw new ExcelEncryptionException(
                    ExcelEncryptionReason.UnsupportedScheme,
                    $"This workbook uses ECMA-376 standard encryption (EncryptionInfo {major}.{minor}), " +
                    "which is recognized but not yet supported."),
                (2, _) or (3, 1) or (4, 1) => throw new ExcelEncryptionException(
                    ExcelEncryptionReason.UnsupportedScheme,
                    $"This workbook uses RC4 CryptoAPI encryption (EncryptionInfo {major}.{minor}), which is not supported."),
                _ => throw new ExcelEncryptionException(
                    ExcelEncryptionReason.UnsupportedScheme,
                    $"Unrecognized EncryptionInfo version {major}.{minor}."),
            };
        }

        // `options` is threaded through for forward compatibility with Task 4's
        // ExcelReaderOptions.MaxPasswordSpinCount; nothing in this pass reads it yet.
        private static AgileDescriptor ParseAgile(ReadOnlySpan<byte> xml, ExcelReaderOptions options)
        {
            _ = options;
            CryptoParameters? keyData = null;
            CryptoParameters? passwordEncryptor = null;
            byte[] encryptedHmacKey = [];
            byte[] encryptedHmacValue = [];

            // XmlReader needs a Stream/TextReader, not a span; this is the one-shot ~1 KB copy the
            // file header comment above calls out — not on any hot path.
            using var stream = new MemoryStream(xml.ToArray(), writable: false);
            try
            {
                using XmlReader reader = XmlReader.Create(stream, XmlSettings);
                while (reader.Read())
                {
                    if (reader.NodeType != XmlNodeType.Element)
                    {
                        continue;
                    }
                    if (string.Equals(reader.LocalName, "keyData", StringComparison.Ordinal)
                        && (reader.NamespaceURI.Length == 0 || string.Equals(reader.NamespaceURI, EncryptionNamespace, StringComparison.Ordinal)))
                    {
                        keyData = ReadCryptoParameters(reader, options);
                    }
                    else if (string.Equals(reader.LocalName, "dataIntegrity", StringComparison.Ordinal)
                        && (reader.NamespaceURI.Length == 0 || string.Equals(reader.NamespaceURI, EncryptionNamespace, StringComparison.Ordinal)))
                    {
                        encryptedHmacKey = ReadBase64Attribute(reader, "encryptedHmacKey") ?? [];
                        encryptedHmacValue = ReadBase64Attribute(reader, "encryptedHmacValue") ?? [];
                    }
                    else if (string.Equals(reader.LocalName, "encryptedKey", StringComparison.Ordinal) && passwordEncryptor is null
                        && (reader.NamespaceURI.Length == 0 || string.Equals(reader.NamespaceURI, PasswordKeyEncryptorNamespace, StringComparison.Ordinal)))
                    {
                        passwordEncryptor = ReadCryptoParameters(reader, options);
                    }
                }
            }
            catch (XmlException ex)
            {
                throw new InvalidDataException("The EncryptionInfo descriptor is not well-formed XML.", ex);
            }

            if (keyData is null)
            {
                throw new InvalidDataException("The EncryptionInfo descriptor is missing its <keyData> element.");
            }
            if (passwordEncryptor is null)
            {
                // Not malformed — e.g. a certificate-only key encryptor is a legitimate agile
                // descriptor shape, just one this library cannot open without a password to check.
                throw new ExcelEncryptionException(ExcelEncryptionReason.UnsupportedScheme,
                    "The EncryptionInfo descriptor has no password key encryptor.");
            }

            return new AgileDescriptor(keyData, passwordEncryptor, encryptedHmacKey, encryptedHmacValue);
        }

        // Reads and validates the attribute set shared by <keyData> and <p:encryptedKey>. keyData
        // never carries spinCount/encryptedVerifierHash*/encryptedKeyValue, so those default to
        // 0/empty for it — callers only read the fields their element actually populates.
        private static CryptoParameters ReadCryptoParameters(XmlReader reader, ExcelReaderOptions options)
        {
            int saltSize = ReadIntAttribute(reader, "saltSize");
            int blockSize = ReadIntAttribute(reader, "blockSize");
            int keyBits = ReadIntAttribute(reader, "keyBits");
            int hashSize = ReadIntAttribute(reader, "hashSize");
            HashKind hash = ParseHash(reader.GetAttribute("hashAlgorithm"));
            string? cipherAlgorithm = reader.GetAttribute("cipherAlgorithm");
            string? cipherChaining = reader.GetAttribute("cipherChaining");
            byte[] saltValue = ReadBase64Attribute(reader, "saltValue") ?? [];
            int spinCount = ReadIntAttribute(reader, "spinCount");
            byte[] encryptedVerifierHashInput = ReadBase64Attribute(reader, "encryptedVerifierHashInput") ?? [];
            byte[] encryptedVerifierHashValue = ReadBase64Attribute(reader, "encryptedVerifierHashValue") ?? [];
            byte[] encryptedKeyValue = ReadBase64Attribute(reader, "encryptedKeyValue") ?? [];

            if (cipherAlgorithm is null || !cipherAlgorithm.StartsWith("AES", StringComparison.Ordinal))
            {
                throw new ExcelEncryptionException(ExcelEncryptionReason.UnsupportedScheme,
                    $"Unsupported cipher algorithm '{cipherAlgorithm}' in the encryption descriptor.");
            }
            // The specification permits CFB chaining; Excel never writes it and no fixture can be
            // obtained to verify a decrypt against it, so it is rejected rather than half-supported.
            if (!string.Equals(cipherChaining, "ChainingModeCBC", StringComparison.Ordinal))
            {
                throw new ExcelEncryptionException(ExcelEncryptionReason.UnsupportedScheme,
                    $"Unsupported cipher chaining mode '{cipherChaining}' in the encryption descriptor.");
            }
            if (keyBits is not (128 or 192 or 256))
            {
                throw new ExcelEncryptionException(ExcelEncryptionReason.UnsupportedScheme,
                    $"Unsupported key size ({keyBits} bits) in the encryption descriptor.");
            }
            if (blockSize is < 1 or > 64)
            {
                throw new InvalidDataException($"The encryption descriptor's block size ({blockSize}) is out of range.");
            }
            if (saltSize is < 1 or > 64 || saltSize != saltValue.Length)
            {
                throw new InvalidDataException("The encryption descriptor's salt size does not match its salt value.");
            }
            if (hashSize is < 1 or > 64)
            {
                throw new InvalidDataException($"The encryption descriptor's hash size ({hashSize}) is out of range.");
            }
            // A resource limit, not a scheme problem, so this throws ExcelLimitExceededException
            // rather than ExcelEncryptionException — a huge spin count is a DoS knob, not evidence the
            // file uses a scheme this library can't parse.
            if (spinCount < 0 || spinCount > options.MaxPasswordSpinCount)
            {
                throw new ExcelLimitExceededException(nameof(options.MaxPasswordSpinCount), options.MaxPasswordSpinCount, spinCount);
            }

            return new CryptoParameters(saltSize, blockSize, keyBits, hashSize, hash, saltValue, spinCount,
                encryptedVerifierHashInput, encryptedVerifierHashValue, encryptedKeyValue);
        }

        private static HashKind ParseHash(string? name)
        {
            return name switch
            {
                "SHA1" => HashKind.Sha1,
                "SHA256" or "SHA-256" => HashKind.Sha256,
                "SHA384" or "SHA-384" => HashKind.Sha384,
                "SHA512" or "SHA-512" => HashKind.Sha512,
                // Never HashAlgorithm.Create(name): reflecting a file-supplied string is both an AOT
                // hazard and an injection surface.
                _ => throw new ExcelEncryptionException(ExcelEncryptionReason.UnsupportedScheme,
                        $"Unsupported hash algorithm '{name}' in the encryption descriptor."),
            };
        }

        private static int ReadIntAttribute(XmlReader reader, string name)
        {
            string? value = reader.GetAttribute(name);
            if (value is null)
            {
                return 0;
            }
            if (!int.TryParse(value, NumberStyles.None, CultureInfo.InvariantCulture, out int result))
            {
                throw new InvalidDataException($"The encryption descriptor's '{name}' attribute is not a valid integer.");
            }
            return result;
        }

        // Base64-decodes with Convert.TryFromBase64String (never Convert.FromBase64String, which
        // throws FormatException instead of letting a malformed attribute surface as the
        // InvalidDataException every other bound in this parser throws).
        private static byte[]? ReadBase64Attribute(XmlReader reader, string name)
        {
            string? value = reader.GetAttribute(name);
            if (value is null)
            {
                return null;
            }
            byte[] buffer = new byte[(value.Length / 4 * 3) + 3];
            if (!Convert.TryFromBase64String(value, buffer, out int written))
            {
                throw new InvalidDataException($"The encryption descriptor's '{name}' attribute is not valid base64.");
            }
            return buffer[..written];
        }
    }

    // Populated from <keyData> and from <p:encryptedKey> under the password <keyEncryptor>; both
    // elements share this attribute shape. keyData never carries SpinCount / the encrypted* fields —
    // they default to 0 / empty for it. See EncryptionDescriptor.ReadCryptoParameters.
    internal sealed record CryptoParameters(
        int SaltSize,
        int BlockSize,
        int KeyBits,
        int HashSize,
        HashKind Hash,
        byte[] SaltValue,
        int SpinCount,
        byte[] EncryptedVerifierHashInput,
        byte[] EncryptedVerifierHashValue,
        byte[] EncryptedKeyValue);

    internal sealed record AgileDescriptor(
        CryptoParameters KeyData,
        CryptoParameters PasswordEncryptor,
        byte[] EncryptedHmacKey,
        byte[] EncryptedHmacValue) : EncryptionDescriptor
    {
        // ParseAgile defaults both fields to [] when the descriptor has no <dataIntegrity> element at
        // all (older writers can omit it) — that's the one shape distinguishable from "present but
        // empty", since a real element always carries both attributes. Callers use this to make
        // opting into HMAC verification a documented no-op rather than a spurious failure for such a
        // file (see PackageIntegrity/DecryptedPackageStream.Create/EncryptedPackageOpener.DecryptToMemory).
        internal bool HasDataIntegrity => EncryptedHmacKey.Length > 0 && EncryptedHmacValue.Length > 0;
    }

    internal enum HashKind
    {
        Sha1,
        Sha256,
        Sha384,
        Sha512,
    }
}
