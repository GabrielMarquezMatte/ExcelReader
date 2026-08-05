using System.Numerics;
using System.Runtime.InteropServices;
using System.Runtime.Intrinsics;

namespace ExcelReader.Core.Reader
{
    // Single-pass scanner reporting the position of every CSV control byte (delimiter, quote, CR, LF)
    // in the read buffer, in order. One vector load covers 32 (AVX2) or 16 (SSE2/NEON) bytes; every
    // control byte inside that window afterwards costs only a trailing-zero-count and a mask clear,
    // so a whole quote-free record is located with ceil(length / 32) loads instead of one IndexOf
    // call per field. Bounded strictly by `_len`: the buffer may be a caller-owned array whose logical
    // length is its physical length, so a vector load must never straddle it.
    //
    // A plain (non-ref) struct, backed by `byte[]` rather than a `ReadOnlySpan<byte>`, so
    // CsvReader.Enumerator can hold one instance as a field and persist it across records within the
    // same buffered window instead of reconstructing it (and re-loading vectors over bytes already
    // scanned) per record — see Reset/Continue below. Every method that touches the buffer builds a
    // local ReadOnlySpan<byte> from the `byte[]` field; only the field itself couldn't be a span.
    internal struct CsvControlScanner
    {
        private const byte Cr = (byte)'\r';
        private const byte Lf = (byte)'\n';

        private byte[] _buf;
        private int _len;
        private readonly byte _delim;
        private readonly byte _quote;
        // Absolute offset the bits of _mask are relative to; bit i means a control byte at _chunkStart + i.
        private int _chunkStart;
        private uint _mask;
        // Next byte not yet covered by a loaded chunk.
        private int _pos;

        internal CsvControlScanner(byte delim, byte quote)
        {
            _delim = delim;
            _quote = quote;
            _buf = [];
            _len = 0;
            _pos = 0;
            _chunkStart = 0;
            _mask = 0;
        }

        // Re-anchors the scanner to scan buf[pos..len) from scratch, discarding any pending mask —
        // required whenever the caller can't guarantee the previous scan's leftover state still
        // applies (the buffer was refilled/compacted/grown, or a different parse path advanced the
        // cursor without going through this scanner at all). See CsvReader.Enumerator's _scannerValid.
        internal void Reset(byte[] buf, int len, int pos)
        {
            _buf = buf;
            _len = len;
            _pos = pos;
            _chunkStart = 0;
            _mask = 0;
        }

        // Refreshes the buffer reference/length for a call that continues from exactly where the
        // previous Next()/SkipByte call left off (no Fill happened in between, so _pos/_mask/_chunkStart
        // are still valid) — assigning buf/len again here is cheap and keeps the contract explicit
        // rather than relying on the caller to never need it.
        internal void Continue(byte[] buf, int len)
        {
            _buf = buf;
            _len = len;
        }

        // Absolute index of the next control byte, or -1 once `_len` is reached.
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

        // Reconciles the scanner's state when the caller consumes one extra byte immediately after
        // the last Next() hit without querying the scanner for it — specifically, the LF half of a
        // CRLF terminator, whose CR half was just returned by Next(). Keeps a later Next() call from
        // re-reporting a byte the caller has already accounted for.
        internal void SkipByte(int position)
        {
            if (_mask != 0 && _chunkStart + BitOperations.TrailingZeroCount(_mask) == position)
            {
                // The LF immediately follows the CR's own bit; since bits are only ever consumed in
                // ascending order, if it was already loaded into this chunk's mask it is necessarily
                // the very next one.
                _mask &= _mask - 1;
            }
            if (_pos <= position)
            {
                _pos = position + 1;
            }
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
            ReadOnlySpan<byte> buf = _buf;
            ref byte origin = ref MemoryMarshal.GetReference(buf);
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
            ReadOnlySpan<byte> buf = _buf;
            ref byte origin = ref MemoryMarshal.GetReference(buf);
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
            byte[] buf = _buf;
            while (_pos < _len)
            {
                byte b = buf[_pos];
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
