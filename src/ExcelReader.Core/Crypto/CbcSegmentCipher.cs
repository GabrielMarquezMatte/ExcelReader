using System.Buffers;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;
using System.Security.Cryptography;

namespace ExcelReader.Core.Crypto
{
    internal sealed class CbcSegmentCipher : IDisposable
    {
        internal const int BlockSize = 16;

        private readonly Aes _aes;
        private readonly ICryptoTransform _transform;
        private readonly byte[] _tail = new byte[BlockSize];
        private bool _disposed;

        [SuppressMessage("Security", "CA5401:Do not use CreateEncryptor with non-default IV",
            Justification = "The all-zero IV here is a seed for the transform's chaining state, not the IV any " +
                "segment is encrypted under: ECMA-376 fixes each segment's IV from the keyData salt and the " +
                "segment index, and Encrypt applies it per segment. No two segments share an effective IV.")]
        private CbcSegmentCipher(byte[] key, bool encrypting)
        {
            Aes aes = Aes.Create();
            try
            {
                aes.Mode = CipherMode.CBC;
                aes.Padding = PaddingMode.None;
                _transform = encrypting ? aes.CreateEncryptor(key, _tail) : aes.CreateDecryptor(key, _tail);
                _aes = aes;
            }
            catch
            {
                aes.Dispose();
                throw;
            }
        }

        internal static CbcSegmentCipher CreateDecryptor(byte[] key)
        {
            return new CbcSegmentCipher(key, encrypting: false);
        }

        internal static CbcSegmentCipher CreateEncryptor(byte[] key)
        {
            return new CbcSegmentCipher(key, encrypting: true);
        }

        internal void Decrypt(ReadOnlyMemory<byte> cipher, ReadOnlySpan<byte> iv, Memory<byte> plain)
        {
            Transform(cipher, plain);
            Span<byte> first = plain.Span[..BlockSize];
            XorInto(_tail, first);
            XorInto(iv, first);
            cipher.Span[^BlockSize..].CopyTo(_tail);
        }

        internal void Encrypt(Memory<byte> plain, ReadOnlySpan<byte> iv, Memory<byte> cipher)
        {
            Span<byte> first = plain.Span[..BlockSize];
            XorInto(iv, first);
            XorInto(_tail, first);
            Transform(plain, cipher);
            cipher.Span[^BlockSize..].CopyTo(_tail);
        }

        private void Transform(ReadOnlyMemory<byte> source, Memory<byte> destination)
        {
            if (MemoryMarshal.TryGetArray(source, out ArraySegment<byte> src)
                && MemoryMarshal.TryGetArray((ReadOnlyMemory<byte>)destination, out ArraySegment<byte> dst))
            {
                RequireWholeBlock(_transform.TransformBlock(src.Array!, src.Offset, src.Count, dst.Array!, dst.Offset), src.Count);
                return;
            }
            byte[] staged = ArrayPool<byte>.Shared.Rent(source.Length);
            try
            {
                source.Span.CopyTo(staged);
                RequireWholeBlock(_transform.TransformBlock(staged, 0, source.Length, staged, 0), source.Length);
                staged.AsSpan(0, source.Length).CopyTo(destination.Span);
            }
            finally
            {
                CryptographicOperations.ZeroMemory(staged.AsSpan(0, source.Length));
                ArrayPool<byte>.Shared.Return(staged);
            }
        }

        private static void RequireWholeBlock(int written, int expected)
        {
            if (written != expected)
            {
                throw new CryptographicException(
                    $"The CBC transform consumed {expected} bytes but produced {written}; a segment cipher must transform a whole segment.");
            }
        }

        private static void XorInto(ReadOnlySpan<byte> value, Span<byte> destination)
        {
            for (int i = 0; i < destination.Length; i++)
            {
                destination[i] ^= value[i];
            }
        }

        public void Dispose()
        {
            if (_disposed)
            {
                return;
            }
            _disposed = true;
            CryptographicOperations.ZeroMemory(_tail);
            _transform.Dispose();
            _aes.Dispose();
        }
    }
}
