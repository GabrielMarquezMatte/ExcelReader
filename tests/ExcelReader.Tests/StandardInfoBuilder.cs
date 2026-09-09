using System.Buffers.Binary;
using System.Text;

namespace ExcelReader.Tests
{
    // Forges a standard EncryptionInfo stream. Every field is a parameter because the negative
    // tests exist to corrupt exactly one of them at a time.
    internal static class StandardInfoBuilder
    {
        internal static byte[] Build(
            int major = 4,
            int minor = 2,
            int algId = 0x00006610,
            int algIdHash = 0x00008004,
            int keySize = 256,
            int saltSize = 16,
            int verifierHashSize = 20,
            int saltBytes = 16,
            int verifierHashBytes = 32,
            int? headerSizeOverride = null)
        {
            byte[] cspName = Encoding.Unicode.GetBytes(
                "Microsoft Enhanced RSA and AES Cryptographic Provider\0");
            int headerSize = headerSizeOverride ?? (32 + cspName.Length);

            using var stream = new MemoryStream();
            byte[] word = new byte[4];

            void WriteI32(int value)
            {
                BinaryPrimitives.WriteInt32LittleEndian(word, value);
                stream.Write(word);
            }

            BinaryPrimitives.WriteUInt16LittleEndian(word, (ushort)major);
            BinaryPrimitives.WriteUInt16LittleEndian(word.AsSpan(2), (ushort)minor);
            stream.Write(word);

            WriteI32(0x00000024);
            WriteI32(headerSize);
            WriteI32(0x00000024);
            WriteI32(0);
            WriteI32(algId);
            WriteI32(algIdHash);
            WriteI32(keySize);
            WriteI32(0x00000018);
            WriteI32(0);
            WriteI32(0);
            stream.Write(cspName);

            WriteI32(saltSize);
            stream.Write(new byte[saltBytes]);
            stream.Write(new byte[16]);
            WriteI32(verifierHashSize);
            stream.Write(new byte[verifierHashBytes]);

            return stream.ToArray();
        }
    }
}
