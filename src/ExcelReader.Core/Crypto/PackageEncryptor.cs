using System.Buffers;
using System.Buffers.Binary;
using System.Security.Cryptography;
using ExcelReader.Core.Reader;
using ExcelReader.Core.Writer.Internal;

namespace ExcelReader.Core.Crypto
{
    // Wraps a plaintext OOXML package in an agile-encrypted CFB container: the write direction of
    // EncryptedPackageOpener.
    //
    // The ciphertext is produced twice. The CFB layout needs every stream's size up front and the
    // dataIntegrity HMAC needs the finished ciphertext, so pass 1 encrypts to feed the HMAC and
    // discards the output, then pass 2 encrypts again straight into the container. That trades a
    // second AES pass (deterministic: same key, same IVs, same plaintext) for O(1) memory at any
    // workbook size, and it leaves `destination` free to be non-seekable.
    internal static class PackageEncryptor
    {
        private const int SegmentSize = 4096;
        private const int PrefixSize = 8;
        private const int CipherBlockSize = 16;
        private const int SaltSize = 16;
        private const int KeyBits = 256;
        private const int HashSize = 64;
        private const int SpinCount = 100_000;

        internal static void Encrypt(Stream package, Stream destination, ExcelPassword password)
        {
            if (!destination.CanWrite)
            {
                throw new ArgumentException("The destination stream must be writable.", nameof(destination));
            }
            (byte[] encryptionInfo, Session session) = Prepare(package, password);
            try
            {
                long streamLength = PrefixSize + CipherLength(session.PlainLength);
                OleCompoundWriter.Write(destination,
                [
                    ByteStream("EncryptionInfo", encryptionInfo),
                    new CfbStreamSpec("EncryptedPackage", streamLength,
                        stream =>
                        {
                            stream.Write(session.Prefix);
                            Session.Local local = session.BeginPass();
                            try
                            {
                                int read;
                                while ((read = ReadSegment(session.Package, local.Plain)) > 0)
                                {
                                    int cipherLen = local.EncryptSegment(read);
                                    stream.Write(local.Cipher.AsSpan(0, cipherLen));
                                }
                            }
                            finally
                            {
                                local.Dispose();
                            }
                        },
                        static (_, _) => throw new NotSupportedException("Encrypt uses the synchronous body only.")),
                ]);
            }
            finally
            {
                session.Dispose();
            }
        }

        internal static async ValueTask EncryptAsync(Stream package, Stream destination, ExcelPassword password, CancellationToken ct)
        {
            if (!destination.CanWrite)
            {
                throw new ArgumentException("The destination stream must be writable.", nameof(destination));
            }
            (byte[] encryptionInfo, Session session) = Prepare(package, password);
            try
            {
                long streamLength = PrefixSize + CipherLength(session.PlainLength);
                await OleCompoundWriter.WriteAsync(destination,
                [
                    ByteStream("EncryptionInfo", encryptionInfo),
                    new CfbStreamSpec("EncryptedPackage", streamLength,
                        static _ => throw new NotSupportedException("EncryptAsync uses the asynchronous body only."),
                        async (stream, token) =>
                        {
                            await stream.WriteAsync(session.Prefix, token).ConfigureAwait(false);
                            Session.Local local = session.BeginPass();
                            try
                            {
                                int read;
                                while ((read = await ReadSegmentAsync(session.Package, local.Plain, token).ConfigureAwait(false)) > 0)
                                {
                                    int cipherLen = local.EncryptSegment(read);
                                    await stream.WriteAsync(local.Cipher.AsMemory(0, cipherLen), token).ConfigureAwait(false);
                                }
                            }
                            finally
                            {
                                local.Dispose();
                            }
                        }),
                ], ct).ConfigureAwait(false);
            }
            finally
            {
                session.Dispose();
            }
        }

        private static CfbStreamSpec ByteStream(string name, byte[] content)
        {
            return new CfbStreamSpec(
                name, content.Length,
                stream => stream.Write(content),
                async (stream, token) => await stream.WriteAsync(content, token).ConfigureAwait(false));
        }

        private static long CipherLength(long plainLength)
        {
            long fullSegments = plainLength / SegmentSize;
            long tail = plainLength % SegmentSize;
            return (fullSegments * SegmentSize) + (tail == 0 ? 0 : ((tail + CipherBlockSize - 1) / CipherBlockSize) * CipherBlockSize);
        }

        private static int ReadSegment(Stream source, byte[] buffer)
        {
            int total = 0;
            while (total < SegmentSize)
            {
                int read = source.Read(buffer, total, SegmentSize - total);
                if (read == 0)
                {
                    break;
                }
                total += read;
            }
            return total;
        }

        private static async ValueTask<int> ReadSegmentAsync(Stream source, byte[] buffer, CancellationToken ct)
        {
            int total = 0;
            while (total < SegmentSize)
            {
                int read = await source.ReadAsync(buffer.AsMemory(total, SegmentSize - total), ct).ConfigureAwait(false);
                if (read == 0)
                {
                    break;
                }
                total += read;
            }
            return total;
        }

        // Generates every random value, runs pass 1 to get the HMAC, and builds the descriptor. Leaves
        // `package` rewound to where pass 2 must start.
        private static (byte[] EncryptionInfo, Session Session) Prepare(Stream package, ExcelPassword password)
        {
            ArgumentNullException.ThrowIfNull(package);
            ArgumentNullException.ThrowIfNull(password);
            if (!package.CanRead || !package.CanSeek)
            {
                throw new ArgumentException("The package stream must be readable and seekable.", nameof(package));
            }

            long origin = package.Position;
            long plainLength = package.Length - origin;
            if (plainLength <= 0)
            {
                throw new ArgumentException("The package stream is empty.", nameof(package));
            }
            if (plainLength < 4)
            {
                throw new ArgumentException(
                    "The package stream is not an OOXML package (missing the PK\\x03\\x04 signature).", nameof(package));
            }
            if (password.Chars.IsEmpty)
            {
                throw new ArgumentException("The password must not be empty.", nameof(password));
            }
            EnsureOpcSignature(package, origin);

            byte[] keyDataSalt = RandomNumberGenerator.GetBytes(SaltSize);
            byte[] pwSalt = RandomNumberGenerator.GetBytes(SaltSize);
            byte[] packageKey = RandomNumberGenerator.GetBytes(KeyBits / 8);
            byte[] verifierInput = RandomNumberGenerator.GetBytes(SaltSize);
            byte[] hmacKey = RandomNumberGenerator.GetBytes(HashSize);

            CryptoParameters keyData = new(SaltSize, CipherBlockSize, KeyBits, HashSize, HashKind.Sha512,
                keyDataSalt, 0, [], [], []);
            var session = new Session(package, origin, plainLength, keyData, packageKey);
            try
            {
                byte[] hmacValue = session.ComputeHmac(hmacKey);
                (byte[] wrappedHmacKey, byte[] wrappedHmacValue) =
                    AgileKeyDerivation.WrapHmac(keyData, packageKey, hmacKey, hmacValue);

                CryptoParameters pwSeed = new(SaltSize, CipherBlockSize, KeyBits, HashSize, HashKind.Sha512,
                    pwSalt, SpinCount, [], [], []);
                PasswordEncryptorBlobs blobs =
                    AgileKeyDerivation.DeriveWriteBlobs(pwSeed, password.Chars, packageKey, verifierInput);
                CryptoParameters passwordEncryptor = pwSeed with
                {
                    EncryptedVerifierHashInput = blobs.EncryptedVerifierHashInput,
                    EncryptedVerifierHashValue = blobs.EncryptedVerifierHashValue,
                    EncryptedKeyValue = blobs.EncryptedKeyValue,
                };

                byte[] info = EncryptionInfoWriter.Build(keyData, passwordEncryptor, wrappedHmacKey, wrappedHmacValue);
                package.Position = origin;
                return (info, session);
            }
            catch
            {
                session.Dispose();
                throw;
            }
            finally
            {
                CryptographicOperations.ZeroMemory(hmacKey);
                CryptographicOperations.ZeroMemory(verifierInput);
            }
        }

        // Encrypting a non-OPC input would silently produce a container that decrypts to garbage and
        // fails much later inside the ZIP reader.
        private static void EnsureOpcSignature(Stream package, long origin)
        {
            Span<byte> signature = stackalloc byte[4];
            package.Position = origin;
            package.ReadExactly(signature);
            package.Position = origin;
            if (!signature.SequenceEqual<byte>([0x50, 0x4B, 0x03, 0x04]))
            {
                throw new ArgumentException(
                    "The package stream is not an OOXML package (missing the PK\\x03\\x04 signature).", nameof(package));
            }
        }

        // Holds the package key, the descriptor's keyData, and the buffers the segment loops need.
        // Owns the key material and zeroes it on disposal.
        private sealed class Session : IDisposable
        {
            private readonly AgileDescriptor _descriptor;
            private readonly byte[] _packageKey;
            private readonly long _origin;
            private bool _disposed;

            internal Session(Stream package, long origin, long plainLength, CryptoParameters keyData, byte[] packageKey)
            {
                Package = package;
                _origin = origin;
                PlainLength = plainLength;
                _packageKey = packageKey;
                _descriptor = new AgileDescriptor(keyData, keyData, [], []);
                Prefix = new byte[PrefixSize];
                BinaryPrimitives.WriteInt64LittleEndian(Prefix, plainLength);
            }

            internal Stream Package { get; }

            internal long PlainLength { get; }

            internal byte[] Prefix { get; }

            // Pass 1: encrypt every segment into the HMAC and throw the ciphertext away.
            internal byte[] ComputeHmac(byte[] hmacKey)
            {
                using IncrementalHash hmac = PackageIntegrity.CreateHmac(_descriptor.KeyData.Hash, hmacKey);
                hmac.AppendData(Prefix);
                using (Local local = BeginPass())
                {
                    int read;
                    while ((read = ReadSegment(Package, local.Plain)) > 0)
                    {
                        int cipherLen = local.EncryptSegment(read);
                        hmac.AppendData(local.Cipher.AsSpan(0, cipherLen));
                    }
                }
                return hmac.GetHashAndReset();
            }

            internal Local BeginPass()
            {
                Package.Position = _origin;
                return new Local(_descriptor, _packageKey);
            }

            public void Dispose()
            {
                if (_disposed)
                {
                    return;
                }
                _disposed = true;
                CryptographicOperations.ZeroMemory(_packageKey);
            }

            // Per-pass state: the AES instance, the IV hasher, the segment buffers, and the segment
            // counter that drives the per-segment IV.
            internal sealed class Local : IDisposable
            {
                private readonly AgileDescriptor _descriptor;
                private readonly Aes _aes;
                private readonly IncrementalHash _ivHasher;
                private readonly byte[] _iv;
                private int _segmentIndex;

                internal Local(AgileDescriptor descriptor, byte[] packageKey)
                {
                    _descriptor = descriptor;
                    _aes = Aes.Create();
                    _aes.Mode = CipherMode.CBC;
                    _aes.Padding = PaddingMode.None;
                    _aes.Key = packageKey;
                    _ivHasher = AgileKeyDerivation.CreateHasher(descriptor.KeyData.Hash);
                    _iv = new byte[descriptor.KeyData.BlockSize];
                    Plain = ArrayPool<byte>.Shared.Rent(SegmentSize);
                    Cipher = ArrayPool<byte>.Shared.Rent(SegmentSize);
                }

                internal byte[] Plain { get; }

                internal byte[] Cipher { get; }

                // Encrypts `length` plaintext bytes from Plain into Cipher, zero-padding the tail up to
                // the cipher block, and returns the ciphertext length.
                internal int EncryptSegment(int length)
                {
                    int padded = ((length + CipherBlockSize - 1) / CipherBlockSize) * CipherBlockSize;
                    Plain.AsSpan(length, padded - length).Clear();
                    AgileKeyDerivation.SegmentIv(_descriptor, _segmentIndex, _ivHasher, _iv);
                    _aes.EncryptCbc(Plain.AsSpan(0, padded), _iv, Cipher.AsSpan(0, padded), PaddingMode.None);
                    _segmentIndex++;
                    return padded;
                }

                public void Dispose()
                {
                    CryptographicOperations.ZeroMemory(Plain.AsSpan(0, SegmentSize));
                    ArrayPool<byte>.Shared.Return(Plain);
                    ArrayPool<byte>.Shared.Return(Cipher);
                    _ivHasher.Dispose();
                    _aes.Dispose();
                }
            }
        }
    }
}
