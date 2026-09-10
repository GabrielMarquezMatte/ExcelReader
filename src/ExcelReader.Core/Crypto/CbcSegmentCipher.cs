using System.Buffers;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;
using System.Security.Cryptography;

namespace ExcelReader.Core.Crypto
{
    // ECMA-376 agile encrypts every 4096-byte segment under its own IV, derived from the keyData
    // salt and the segment index. The one-shot Aes.DecryptCbc/EncryptCbc APIs build and tear down a
    // cipher object on every call - a CNG key import per segment on Windows. Measured over 25,600
    // segments (a ~100 MB package) that setup was 115.5 ms of a 141.5 ms total: 82%. The segment
    // size is fixed by the spec, so it does not amortize away with a bigger buffer; it scales
    // linearly with file size, on both the read and the write side.
    //
    // One ICryptoTransform is reused across every segment instead. A CBC transform carries its
    // chaining state across TransformBlock calls - that is what lets CryptoStream feed a cipher in
    // arbitrary buffer sizes - so the state it holds entering a segment is the last ciphertext
    // block it processed, call it S, rather than the IV this segment wants. A single 16-byte XOR
    // corrects for that, because only the first block of a segment reads the chaining state:
    //
    //   decrypt: TransformBlock yields P0 = D(C0) ^ S, so P0 ^= S ^ iv leaves D(C0) ^ iv.
    //   encrypt: feeding P0' = P0 ^ iv ^ S yields E(P0' ^ S) = E(P0 ^ iv).
    //
    // Every later block in the segment chains against its own neighbour, which is exactly CBC, so
    // it needs no correction in either direction. The fix-up reads only the state this class
    // tracks and never assumes which segment produced it, so a caller that seeks
    // (DecryptedPackageStream) may request segments in any order.
    internal sealed class CbcSegmentCipher : IDisposable
    {
        internal const int BlockSize = 16;

        private readonly Aes _aes;
        private readonly ICryptoTransform _transform;
        // The transform's chaining state: the last ciphertext block it processed. Seeded from the
        // all-zero IV the transform is created with, so its starting value is known rather than the
        // random IV Aes.Create() hands out.
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
                aes.Key = key;
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

        // Decrypts one segment under `iv`. `cipher` and `plain` are equal-length, a whole number of
        // cipher blocks, and must not overlap - the correction reads `cipher`'s final block after
        // the transform has written `plain`.
        internal void Decrypt(ReadOnlyMemory<byte> cipher, ReadOnlySpan<byte> iv, Memory<byte> plain)
        {
            Transform(cipher, plain);
            Span<byte> first = plain.Span[..BlockSize];
            XorInto(_tail, first);
            XorInto(iv, first);
            cipher.Span[^BlockSize..].CopyTo(_tail);
        }

        // Encrypts one segment under `iv`. The first block of `plain` is XORed in place as part of
        // the correction: every caller refills its plaintext buffer from the package stream before
        // the next segment, so the mutated bytes are never read again.
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
            // TransformBlock has no Span overload. Every caller hands over array-backed memory (a
            // rented ArrayPool buffer, a whole-file byte[]), so the staging path below never runs in
            // practice. It exists because there is no one-shot fallback that keeps _tail honest: a
            // transform's chaining state only advances by feeding the transform, so a segment routed
            // around it would desynchronise every segment after it.
            if (MemoryMarshal.TryGetArray(source, out ArraySegment<byte> src)
                && MemoryMarshal.TryGetArray((ReadOnlyMemory<byte>)destination, out ArraySegment<byte> dst))
            {
                _transform.TransformBlock(src.Array!, src.Offset, src.Count, dst.Array!, dst.Offset);
                return;
            }
            byte[] staged = ArrayPool<byte>.Shared.Rent(source.Length);
            try
            {
                source.Span.CopyTo(staged);
                _transform.TransformBlock(staged, 0, source.Length, staged, 0);
                staged.AsSpan(0, source.Length).CopyTo(destination.Span);
            }
            finally
            {
                CryptographicOperations.ZeroMemory(staged.AsSpan(0, source.Length));
                ArrayPool<byte>.Shared.Return(staged);
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
