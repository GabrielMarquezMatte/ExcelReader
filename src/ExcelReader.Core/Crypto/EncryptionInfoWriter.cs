using System.Buffers.Binary;
using System.Globalization;
using System.Text;

namespace ExcelReader.Core.Crypto
{
    // Serializes the CFB "EncryptionInfo" stream: an 8-byte version tuple plus the agile descriptor
    // XML. Hand-rolled, matching this codebase's XML custom — the reader's XmlReader use is the one
    // documented departure.
    //
    // Element nesting and attribute order below are copied from a real Excel-written fixture
    // (data/encrypted/agile-aes256-sha512.xlsx). EncryptionDescriptor.Parse is far more lenient than
    // Excel — it finds elements by local name and ignores the structure around them — so matching the
    // fixture, not merely satisfying our own parser, is the requirement here.
    internal static class EncryptionInfoWriter
    {
        private const string EncryptionNamespace = "http://schemas.microsoft.com/office/2006/encryption";
        private const string PasswordNamespace = "http://schemas.microsoft.com/office/2006/keyEncryptor/password";
        private const string CertificateNamespace = "http://schemas.microsoft.com/office/2006/keyEncryptor/certificate";

        internal static byte[] Build(CryptoParameters keyData, CryptoParameters passwordEncryptor,
            ReadOnlySpan<byte> encryptedHmacKey, ReadOnlySpan<byte> encryptedHmacValue)
        {
            var xml = new StringBuilder(1400);
            xml.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n");
            xml.Append("<encryption xmlns=\"").Append(EncryptionNamespace)
               .Append("\" xmlns:p=\"").Append(PasswordNamespace)
               .Append("\" xmlns:c=\"").Append(CertificateNamespace).Append("\">");

            xml.Append("<keyData");
            AppendCommonAttributes(xml, keyData);
            xml.Append("/>");

            xml.Append("<dataIntegrity encryptedHmacKey=\"").Append(Convert.ToBase64String(encryptedHmacKey))
               .Append("\" encryptedHmacValue=\"").Append(Convert.ToBase64String(encryptedHmacValue)).Append("\"/>");

            xml.Append("<keyEncryptors><keyEncryptor uri=\"").Append(PasswordNamespace).Append("\">");
            xml.Append("<p:encryptedKey spinCount=\"")
               .Append(passwordEncryptor.SpinCount.ToString(CultureInfo.InvariantCulture)).Append('"');
            AppendCommonAttributes(xml, passwordEncryptor);
            xml.Append(" encryptedVerifierHashInput=\"").Append(Convert.ToBase64String(passwordEncryptor.EncryptedVerifierHashInput))
               .Append("\" encryptedVerifierHashValue=\"").Append(Convert.ToBase64String(passwordEncryptor.EncryptedVerifierHashValue))
               .Append("\" encryptedKeyValue=\"").Append(Convert.ToBase64String(passwordEncryptor.EncryptedKeyValue))
               .Append("\"/>");
            xml.Append("</keyEncryptor></keyEncryptors></encryption>");

            byte[] body = Encoding.UTF8.GetBytes(xml.ToString());
            byte[] info = new byte[8 + body.Length];
            BinaryPrimitives.WriteUInt16LittleEndian(info.AsSpan(0), 4);   // major
            BinaryPrimitives.WriteUInt16LittleEndian(info.AsSpan(2), 4);   // minor
            BinaryPrimitives.WriteUInt32LittleEndian(info.AsSpan(4), 0x40); // reserved flags, as Excel writes
            body.CopyTo(info.AsSpan(8));
            return info;
        }

        // The attribute set <keyData> and <p:encryptedKey> share, in the order Excel emits them.
        private static void AppendCommonAttributes(StringBuilder xml, CryptoParameters p)
        {
            xml.Append(" saltSize=\"").Append(p.SaltSize.ToString(CultureInfo.InvariantCulture))
               .Append("\" blockSize=\"").Append(p.BlockSize.ToString(CultureInfo.InvariantCulture))
               .Append("\" keyBits=\"").Append(p.KeyBits.ToString(CultureInfo.InvariantCulture))
               .Append("\" hashSize=\"").Append(p.HashSize.ToString(CultureInfo.InvariantCulture))
               .Append("\" cipherAlgorithm=\"AES\" cipherChaining=\"ChainingModeCBC\" hashAlgorithm=\"")
               .Append(HashName(p.Hash))
               .Append("\" saltValue=\"").Append(Convert.ToBase64String(p.SaltValue)).Append('"');
        }

        private static string HashName(HashKind kind)
        {
            return kind switch
            {
                HashKind.Sha1 => "SHA1",
                HashKind.Sha256 => "SHA256",
                HashKind.Sha384 => "SHA384",
                HashKind.Sha512 => "SHA512",
                _ => throw new ArgumentOutOfRangeException(nameof(kind), kind, "Unsupported hash algorithm."),
            };
        }
    }
}
