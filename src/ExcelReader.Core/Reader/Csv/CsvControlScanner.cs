using System.Numerics;
using System.Runtime.InteropServices;
using System.Runtime.Intrinsics;

namespace ExcelReader.Core.Reader.Csv
{
    internal struct CsvControlScanner
    {
        private const byte Cr = (byte)'\r';
        private const byte Lf = (byte)'\n';

        private byte[] _buf;
        private int _len;
        private readonly byte _delim;
        private readonly byte _quote;
        private int _chunkStart;
        private uint _mask;
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

        internal void Reset(byte[] buf, int len, int pos)
        {
            _buf = buf;
            _len = len;
            _pos = pos;
            _chunkStart = 0;
            _mask = 0;
        }

        internal void Continue(byte[] buf, int len)
        {
            _buf = buf;
            _len = len;
        }

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

        internal bool TryTakeMask(out uint mask, out int chunkStart)
        {
            mask = _mask;
            chunkStart = _chunkStart;
            if (mask != 0)
            {
                _mask = 0;
                return true;
            }
            if (Vector256.IsHardwareAccelerated)
            {
                return TryLoadVector256(out mask, out chunkStart);
            }
            return Vector128.IsHardwareAccelerated && TryLoadVector128(out mask, out chunkStart);
        }

        internal void PutBack(uint mask, int chunkStart)
        {
            _mask = mask;
            _chunkStart = chunkStart;
        }

        internal void SkipByte(int position)
        {
            if (_mask != 0 && _chunkStart + BitOperations.TrailingZeroCount(_mask) == position)
            {
                _mask &= _mask - 1;
            }
            if (_pos <= position)
            {
                _pos = position + 1;
            }
        }

        private int NextFromChunks()
        {
            if (TryTakeMask(out uint mask, out int chunkStart))
            {
                _chunkStart = chunkStart;
                _mask = mask & (mask - 1);
                return chunkStart + BitOperations.TrailingZeroCount(mask);
            }
            return NextScalar();
        }

        private bool TryLoadVector256(out uint mask, out int chunkStart)
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
                mask = (Vector256.Equals(chunk, delim)
                      | Vector256.Equals(chunk, quote)
                      | Vector256.Equals(chunk, cr)
                      | Vector256.Equals(chunk, lf)).ExtractMostSignificantBits();
                chunkStart = _pos;
                _pos += Vector256<byte>.Count;
                if (mask != 0)
                {
                    return true;
                }
            }
            mask = 0;
            chunkStart = 0;
            return false;
        }

        private bool TryLoadVector128(out uint mask, out int chunkStart)
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
                mask = (Vector128.Equals(chunk, delim)
                      | Vector128.Equals(chunk, quote)
                      | Vector128.Equals(chunk, cr)
                      | Vector128.Equals(chunk, lf)).ExtractMostSignificantBits();
                chunkStart = _pos;
                _pos += Vector128<byte>.Count;
                if (mask != 0)
                {
                    return true;
                }
            }
            mask = 0;
            chunkStart = 0;
            return false;
        }

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
