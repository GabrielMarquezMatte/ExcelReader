using System.Buffers.Binary;
using System.Globalization;
using System.Xml;
using ExcelReader.Core.Reader;

namespace ExcelReader.Core.Crypto
{
    // Parses the CFB "EncryptionInfo" stream: an 8-byte little-endian (major, minor) version tuple,
    // followed by either a UTF-8 XML descriptor (agile, 4.4) or a binary header+verifier (standard,
    // 3.2/4.2). Runs on wholly untrusted input before any password check.
    //
    // ParseAgile uses XmlReader — the one departure from this codebase's hand-rolled XML custom. The
    // descriptor is ~1 KB, parsed once on a cold path, so throughput isn't the concern here.
    internal abstract record EncryptionDescriptor
    {
        private const string EncryptionNamespace = "http://schemas.microsoft.com/office/2006/encryption";
        private const string PasswordKeyEncryptorNamespace = "http://schemas.microsoft.com/office/2006/keyEncryptor/password";

        // Untrusted, pre-authentication input: DTD processing and external resolution stay off.
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
                // The binary descriptor's own fields start right after the 4-byte version tuple.
                // Agile skips 8 instead because a 4-byte reserved field sits between its version
                // and its XML; the asymmetry is in the format, not a mistake here.
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

        // [MS-OFFCRYPTO] 2.3.4.5-2.3.4.6. Versions 2.2/3.2/4.2 all carry this binary
        // header+verifier, for standard AES *and* for RC4 CryptoAPI — the cipher is named by
        // algId, so the header must be read before the scheme can be decided.
        private static StandardDescriptor ParseBinary(ReadOnlySpan<byte> body, int major, int minor)
        {
            // headerFlags, then the header's own byte count.
            if (body.Length < 8)
            {
                throw new InvalidDataException("The EncryptionInfo stream is truncated.");
            }
            int headerSize = BinaryPrimitives.ReadInt32LittleEndian(body[4..]);
            // Eight fixed uint32 fields precede the variable-length CSP name.
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
            // 0 means "the scheme's default", which for standard encryption is SHA-1.
            if (algIdHash is not (0x00008004 or 0))
            {
                throw new ExcelEncryptionException(ExcelEncryptionReason.UnsupportedScheme,
                    $"Unsupported hash algorithm (algIdHash 0x{algIdHash:X4}) in the encryption descriptor.");
            }

            return ParseVerifier(body[(8 + headerSize)..], algId, keyBits);
        }

        // AES key size is implied by algId, so a stated keySize that disagrees with it means the
        // descriptor is internally inconsistent and one of the two would have to be ignored.
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

        // [MS-OFFCRYPTO] 2.3.3 EncryptionVerifier: both sizes are fixed for AES, so a different
        // value is a malformed descriptor rather than a scheme this library declines to support.
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
            int verifierHashSize = BinaryPrimitives.ReadInt32LittleEndian(
                verifier[(4 + SaltLength + VerifierLength)..]);
            if (verifierHashSize != VerifierHashLength)
            {
                throw new InvalidDataException(
                    $"The encryption descriptor declares a {verifierHashSize}-byte verifier hash; AES requires {VerifierHashLength}.");
            }

            byte[] salt = verifier.Slice(4, SaltLength).ToArray();
            byte[] encryptedVerifier = verifier.Slice(4 + SaltLength, VerifierLength).ToArray();
            byte[] encryptedVerifierHash = verifier
                .Slice(8 + SaltLength + VerifierLength, VerifierHashLength).ToArray();

            return new StandardDescriptor(algId, keyBits, salt, encryptedVerifier, encryptedVerifierHash);
        }

        private static AgileDescriptor ParseAgile(ReadOnlySpan<byte> xml, ExcelReaderOptions options)
        {
            _ = options;
            CryptoParameters? keyData = null;
            CryptoParameters? passwordEncryptor = null;
            byte[] encryptedHmacKey = [];
            byte[] encryptedHmacValue = [];

            // XmlReader needs a Stream/TextReader, not a span.
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
                // Not malformed: a certificate-only key encryptor is a legitimate shape this library
                // simply can't open.
                throw new ExcelEncryptionException(ExcelEncryptionReason.UnsupportedScheme,
                    "The EncryptionInfo descriptor has no password key encryptor.");
            }

            return new AgileDescriptor(keyData, passwordEncryptor, encryptedHmacKey, encryptedHmacValue);
        }

        // Reads and validates the attribute set shared by <keyData> and <p:encryptedKey>. keyData
        // never carries spinCount/encryptedVerifierHash*/encryptedKeyValue, which default to 0/empty.
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
            // CFB chaining is spec-permitted but has no fixture to verify a decrypt against.
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
            // AES-CBC has a fixed 16-byte block size; any other declared value is rejected here rather
            // than surfacing as a raw CryptographicException out of `aes.IV = ...`.
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
            // A resource limit (DoS knob), not a scheme problem, so ExcelLimitExceededException here.
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

    // Populated from <keyData> and from <p:encryptedKey>; keyData leaves SpinCount/encrypted* at
    // 0/empty. See EncryptionDescriptor.ReadCryptoParameters.
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
        // False only when the descriptor has no <dataIntegrity> element at all (older writers can omit
        // it) — a real element always carries both attributes.
        internal bool HasDataIntegrity => EncryptedHmacKey.Length > 0 && EncryptedHmacValue.Length > 0;
    }

    // ECMA-376 standard encryption (EncryptionInfo 3.2/4.2). Deliberately not reusing
    // CryptoParameters: its SpinCount/EncryptedKeyValue/HashSize/BlockSize fields describe agile's
    // XML descriptor and would sit permanently empty here, implying they mean something. The hash
    // is not a field either — standard encryption is SHA-1 by definition, and algIdHash is
    // validated at parse time rather than carried forward.
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
