using System.Buffers.Binary;
using System.Globalization;
using System.Xml;
using ExcelReader.Core.Reader;

namespace ExcelReader.Core.Crypto
{
    // Parses the CFB "EncryptionInfo" stream: an 8-byte little-endian (major, minor) version tuple,
    internal abstract record EncryptionDescriptor
    {
        private const string EncryptionNamespace = "http://schemas.microsoft.com/office/2006/encryption";
        private const string PasswordKeyEncryptorNamespace = "http://schemas.microsoft.com/office/2006/keyEncryptor/password";

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
                (2, 2) or (3, 2) or (4, 2) => ParseBinary(info[4..], major, minor),
                (3, 1) or (4, 1) => throw new ExcelEncryptionException(
                    ExcelEncryptionReason.UnsupportedScheme,
                    $"This workbook uses RC4 CryptoAPI encryption (EncryptionInfo {major}.{minor}), which is not supported."),
                (3, 3) or (4, 3) => throw new ExcelEncryptionException(
                    ExcelEncryptionReason.UnsupportedScheme,
                    $"This workbook uses extensible encryption (EncryptionInfo {major}.{minor}), which is not supported."),
                _ => throw new ExcelEncryptionException(
                    ExcelEncryptionReason.UnsupportedScheme,
                    $"Unrecognized EncryptionInfo version {major}.{minor}."),
            };
        }

        private static StandardDescriptor ParseBinary(ReadOnlySpan<byte> body, int major, int minor)
        {
            if (body.Length < 8)
            {
                throw new InvalidDataException("The EncryptionInfo stream is truncated.");
            }
            int headerSize = BinaryPrimitives.ReadInt32LittleEndian(body[4..]);
            if (headerSize < 32 || headerSize > body.Length - 8)
            {
                throw new InvalidDataException(
                    "The EncryptionInfo descriptor's header size does not fit its stream.");
            }

            ReadOnlySpan<byte> header = body.Slice(8, headerSize);
            int algId = BinaryPrimitives.ReadInt32LittleEndian(header[8..]);
            int algIdHash = BinaryPrimitives.ReadInt32LittleEndian(header[12..]);
            int keySize = BinaryPrimitives.ReadInt32LittleEndian(header[16..]);

            int keyBits = ResolveKeyBits(algId, keySize, major, minor);
            if (algIdHash is not (0x00008004 or 0))
            {
                throw new ExcelEncryptionException(ExcelEncryptionReason.UnsupportedScheme,
                    $"Unsupported hash algorithm (algIdHash 0x{algIdHash:X4}) in the encryption descriptor.");
            }

            return ParseVerifier(body[(8 + headerSize)..], algId, keyBits);
        }

        private static int ResolveKeyBits(int algId, int keySize, int major, int minor)
        {
            int implied = algId switch
            {
                0x0000660E => 128,
                0x0000660F => 192,
                0x00006610 => 256,
                0x00006801 => throw new ExcelEncryptionException(
                    ExcelEncryptionReason.UnsupportedScheme,
                    $"This workbook uses RC4 CryptoAPI encryption (EncryptionInfo {major}.{minor}), which is not supported."),
                _ => throw new ExcelEncryptionException(
                    ExcelEncryptionReason.UnsupportedScheme,
                    $"Unsupported cipher algorithm (algId 0x{algId:X4}) in the encryption descriptor."),
            };
            if (keySize != 0 && keySize != implied)
            {
                throw new ExcelEncryptionException(ExcelEncryptionReason.UnsupportedScheme,
                    $"The encryption descriptor's key size ({keySize} bits) disagrees with its cipher algorithm.");
            }
            return implied;
        }

        private static StandardDescriptor ParseVerifier(ReadOnlySpan<byte> verifier, int algId, int keyBits)
        {
            const int SaltLength = 16;
            const int VerifierLength = 16;
            const int VerifierHashLength = 32;

            if (verifier.Length < 8 + SaltLength + VerifierLength + VerifierHashLength)
            {
                throw new InvalidDataException("The EncryptionInfo descriptor's verifier is truncated.");
            }
            int saltSize = BinaryPrimitives.ReadInt32LittleEndian(verifier);
            if (saltSize != SaltLength)
            {
                throw new InvalidDataException(
                    $"The encryption descriptor declares a {saltSize}-byte salt; AES requires {SaltLength}.");
            }

            byte[] salt = verifier.Slice(4, SaltLength).ToArray();
            byte[] encryptedVerifier = verifier.Slice(4 + SaltLength, VerifierLength).ToArray();
            byte[] encryptedVerifierHash = verifier
                .Slice(8 + SaltLength + VerifierLength, VerifierHashLength).ToArray();

            return new StandardDescriptor(algId, keyBits, salt, encryptedVerifier, encryptedVerifierHash);
        }

        private static AgileDescriptor ParseAgile(ReadOnlySpan<byte> xml, ExcelReaderOptions options)
        {
            CryptoParameters? keyData = null;
            CryptoParameters? passwordEncryptor = null;
            byte[] encryptedHmacKey = [];
            byte[] encryptedHmacValue = [];

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
                throw new ExcelEncryptionException(ExcelEncryptionReason.UnsupportedScheme,
                    "The EncryptionInfo descriptor has no password key encryptor.");
            }

            return new AgileDescriptor(keyData, passwordEncryptor, encryptedHmacKey, encryptedHmacValue);
        }

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
            if (blockSize != 16)
            {
                throw new ExcelEncryptionException(ExcelEncryptionReason.UnsupportedScheme,
                    $"Unsupported block size ({blockSize}) in the encryption descriptor; AES-CBC requires 16.");
            }
            if (saltSize is < 1 or > 64 || saltSize != saltValue.Length)
            {
                throw new InvalidDataException("The encryption descriptor's salt size does not match its salt value.");
            }
            if (hashSize is < 1 or > 64)
            {
                throw new InvalidDataException($"The encryption descriptor's hash size ({hashSize}) is out of range.");
            }
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
        internal bool HasDataIntegrity => EncryptedHmacKey.Length > 0 && EncryptedHmacValue.Length > 0;
    }

    internal sealed record StandardDescriptor(
        int AlgId,
        int KeyBits,
        byte[] Salt,
        byte[] EncryptedVerifier,
        byte[] EncryptedVerifierHash) : EncryptionDescriptor;

    internal enum HashKind
    {
        Sha1,
        Sha256,
        Sha384,
        Sha512,
    }
}
