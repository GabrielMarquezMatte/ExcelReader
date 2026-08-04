using System.Numerics;
using System.Runtime.InteropServices;
using System.Runtime.Intrinsics;

namespace ExcelReader.Core.Reader
{
    // Single-pass scanner reporting the position of every CSV control byte (delimiter, quote, CR, LF)
    // in the read buffer, in order. One vector load covers 32 (AVX2) or 16 (SSE2/NEON) bytes; every
    // control byte inside that window afterwards costs only a trailing-zero-count and a mask clear,
    // so a whole quote-free record is located with ceil(length / 32) loads instead of one IndexOf
    // call per field. Bounded strictly by `len`: the buffer may be a caller-owned array whose logical
    // length is its physical length, so a vector load must never straddle it.
    internal ref struct CsvControlScanner
    {
        private const byte Cr = (byte)'\r';
        private const byte Lf = (byte)'\n';

        private readonly ReadOnlySpan<byte> _buf;
        private readonly int _len;
        private readonly byte _delim;
        private readonly byte _quote;
        // Absolute offset the bits of _mask are relative to; bit i means a control byte at _chunkStart + i.
        private int _chunkStart;
        private uint _mask;
        // Next byte not yet covered by a loaded chunk.
        private int _pos;

        internal CsvControlScanner(ReadOnlySpan<byte> buf, int start, int len, byte delim, byte quote)
        {
            _buf = buf;
            _len = len;
            _delim = delim;
            _quote = quote;
            _pos = start;
            _chunkStart = 0;
            _mask = 0;
        }

        // Absolute index of the next control byte, or -1 once `len` is reached.
        internal int Next()
        {
            if (_mask != 0)
            {
                int bit = BitOperations.TrailingZeroCount(_mask);
                _mask &= _mask - 1;
                return _chunkStart + bit;
            }
            return NextFromChunks();
        }

        // IsHardwareAccelerated is a JIT-time constant, so only one of these branches survives codegen.
        private int NextFromChunks()
        {
            if (Vector256.IsHardwareAccelerated)
            {
                int found = NextVector256();
                if (found >= 0)
                {
                    return found;
                }
            }
            else if (Vector128.IsHardwareAccelerated)
            {
                int found = NextVector128();
                if (found >= 0)
                {
                    return found;
                }
            }
            return NextScalar();
        }

        private int NextVector256()
        {
            Vector256<byte> delim = Vector256.Create(_delim);
            Vector256<byte> quote = Vector256.Create(_quote);
            Vector256<byte> cr = Vector256.Create(Cr);
            Vector256<byte> lf = Vector256.Create(Lf);
            ref byte origin = ref MemoryMarshal.GetReference(_buf);
            while (_pos + Vector256<byte>.Count <= _len)
            {
                Vector256<byte> chunk = Vector256.LoadUnsafe(ref origin, (nuint)_pos);
                uint mask = (Vector256.Equals(chunk, delim)
                           | Vector256.Equals(chunk, quote)
                           | Vector256.Equals(chunk, cr)
                           | Vector256.Equals(chunk, lf)).ExtractMostSignificantBits();
                int start = _pos;
                _pos += Vector256<byte>.Count;
                if (mask == 0)
                {
                    continue;
                }
                _chunkStart = start;
                int bit = BitOperations.TrailingZeroCount(mask);
                _mask = mask & (mask - 1);
                return start + bit;
            }
            return -1;
        }

        // Twin of NextVector256 over 16-byte chunks. Kept as a separate method rather than generic over
        // the vector width: the JIT specializes each one against its own IsHardwareAccelerated constant.
        private int NextVector128()
        {
            Vector128<byte> delim = Vector128.Create(_delim);
            Vector128<byte> quote = Vector128.Create(_quote);
            Vector128<byte> cr = Vector128.Create(Cr);
            Vector128<byte> lf = Vector128.Create(Lf);
            ref byte origin = ref MemoryMarshal.GetReference(_buf);
            while (_pos + Vector128<byte>.Count <= _len)
            {
                Vector128<byte> chunk = Vector128.LoadUnsafe(ref origin, (nuint)_pos);
                uint mask = (Vector128.Equals(chunk, delim)
                           | Vector128.Equals(chunk, quote)
                           | Vector128.Equals(chunk, cr)
                           | Vector128.Equals(chunk, lf)).ExtractMostSignificantBits();
                int start = _pos;
                _pos += Vector128<byte>.Count;
                if (mask == 0)
                {
                    continue;
                }
                _chunkStart = start;
                int bit = BitOperations.TrailingZeroCount(mask);
                _mask = mask & (mask - 1);
                return start + bit;
            }
            return -1;
        }

        // Handles the sub-chunk tail, and is the whole scan when no SIMD is available.
        private int NextScalar()
        {
            while (_pos < _len)
            {
                byte b = _buf[_pos];
                if (b == _delim || b == _quote || b == Cr || b == Lf)
                {
                    int found = _pos;
                    _pos++;
                    return found;
                }
                _pos++;
            }
            return -1;
        }
    }
}
