using System.Numerics;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.Intrinsics;
using System.Runtime.Intrinsics.X86;

namespace ExcelReader.Core.Reader.Csv
{
    /// <summary>
    /// Classifies 64-byte blocks of a CSV buffer into field ends and escaped-quote marks,
    /// with quoted regions masked out by a prefix-XOR over the quote bits.
    /// </summary>
    /// <remarks>
    /// Only RFC 4180-clean input is classified: a quote that does not open a field or text after a
    /// closing quote makes the scanner stop at that byte and report
    /// <see cref="Blocked"/>, leaving the rest of the record to the reader's generic path.
    /// Scanning must start at a record start; the carried state is meaningless anywhere else.
    /// </remarks>
    internal struct CsvStructuralScanner
    {
        private const byte Cr = (byte)'\r';
        private const byte Lf = (byte)'\n';
        private const int BlockSize = 64;

        private byte[] _buf;
        private int _len;
        private readonly byte _delim;
        private readonly byte _quote;
        private int _pos;
        private bool _blocked;

        private bool _hasPending;
        private ulong _ends;
        private ulong _escapes;
        private int _blockStart;

        private ulong _insideQuotes;
        private ulong _prevStructural;
        private ulong _prevClosing;
        private ulong _prevCr;

        internal CsvStructuralScanner(byte delim, byte quote)
        {
            _delim = delim;
            _quote = quote;
            _buf = [];
        }

        internal readonly bool Blocked
        {
            get
            {
                return _blocked;
            }
        }

        internal readonly bool EndsInsideQuotes
        {
            get
            {
                return _insideQuotes != 0;
            }
        }

        internal static bool Supports(byte delim, byte quote)
        {
            return delim != quote && delim != Cr && delim != Lf && quote != Cr && quote != Lf;
        }

        internal void Reset(byte[] buf, int len, int recordStart)
        {
            _buf = buf;
            _len = len;
            _pos = recordStart;
            _blocked = false;
            _hasPending = false;
            _insideQuotes = 0;
            _prevStructural = 1;
            _prevClosing = 0;
            _prevCr = 0;
        }

        /// <summary>Field-end bits of the block taken last; a set bit is a delimiter or record terminator; for CRLF only the CR is set.</summary>
        internal readonly ulong Ends
        {
            get
            {
                return _ends;
            }
        }

        /// <summary>Bits on the second quote of each doubled quote inside a quoted field.</summary>
        internal readonly ulong Escapes
        {
            get
            {
                return _escapes;
            }
        }

        internal readonly int BlockStart
        {
            get
            {
                return _blockStart;
            }
        }

        internal void PutBack(ulong ends, ulong escapes)
        {
            _hasPending = true;
            _ends = ends;
            _escapes = escapes;
        }

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        internal bool TryTakeBlock()
        {
            if (_hasPending)
            {
                _hasPending = false;
                return true;
            }
            return TryLoadBlock();
        }

        [MethodImpl(MethodImplOptions.NoInlining)]
        private bool TryLoadBlock()
        {
            if (_blocked || _pos >= _len)
            {
                return false;
            }
            ulong quotes;
            ulong delims;
            ulong crs;
            ulong lfs;
            ulong valid = ulong.MaxValue;
            if (_pos + BlockSize <= _len)
            {
                ref byte block = ref Unsafe.Add(ref MemoryMarshal.GetArrayDataReference(_buf), _pos);
                LoadMasksAt(ref block, out quotes, out delims, out crs, out lfs);
            }
            else
            {
                valid = LoadTailMasks(out quotes, out delims, out crs, out lfs);
            }
            _blockStart = _pos;
            _pos += BlockSize;
            if ((quotes | _insideQuotes | _prevClosing) == 0)
            {
                ClassifyUnquoted(delims, crs, lfs);
            }
            else
            {
                Classify(quotes, delims, crs, lfs, valid);
            }
            return true;
        }

        [MethodImpl(MethodImplOptions.NoInlining)]
        private readonly ulong LoadTailMasks(out ulong quotes, out ulong delims, out ulong crs, out ulong lfs)
        {
            Span<byte> tail = stackalloc byte[BlockSize];
            tail.Clear();
            _buf.AsSpan(_pos, _len - _pos).CopyTo(tail);
            ulong valid = (1UL << (_len - _pos)) - 1;
            LoadMasksAt(ref MemoryMarshal.GetReference(tail), out quotes, out delims, out crs, out lfs);
            quotes &= valid;
            delims &= valid;
            crs &= valid;
            lfs &= valid;
            return valid;
        }

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        private readonly void LoadMasksAt(ref byte block, out ulong quotes, out ulong delims, out ulong crs, out ulong lfs)
        {
            if (Vector256.IsHardwareAccelerated)
            {
                Vector256<byte> lo = Vector256.LoadUnsafe(ref block);
                Vector256<byte> hi = Vector256.LoadUnsafe(ref block, 32);
                quotes = Matches(lo, hi, _quote);
                delims = Matches(lo, hi, _delim);
                crs = Matches(lo, hi, Cr);
                lfs = Matches(lo, hi, Lf);
                return;
            }
            Vector128<byte> v0 = Vector128.LoadUnsafe(ref block);
            Vector128<byte> v1 = Vector128.LoadUnsafe(ref block, 16);
            Vector128<byte> v2 = Vector128.LoadUnsafe(ref block, 32);
            Vector128<byte> v3 = Vector128.LoadUnsafe(ref block, 48);
            quotes = Matches(v0, v1, v2, v3, _quote);
            delims = Matches(v0, v1, v2, v3, _delim);
            crs = Matches(v0, v1, v2, v3, Cr);
            lfs = Matches(v0, v1, v2, v3, Lf);
        }

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        private static ulong Matches(Vector256<byte> lo, Vector256<byte> hi, byte value)
        {
            Vector256<byte> target = Vector256.Create(value);
            ulong low = Vector256.Equals(lo, target).ExtractMostSignificantBits();
            ulong high = Vector256.Equals(hi, target).ExtractMostSignificantBits();
            return low | (high << 32);
        }

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        private static ulong Matches(Vector128<byte> v0, Vector128<byte> v1, Vector128<byte> v2, Vector128<byte> v3, byte value)
        {
            Vector128<byte> target = Vector128.Create(value);
            ulong m0 = Vector128.Equals(v0, target).ExtractMostSignificantBits();
            ulong m1 = Vector128.Equals(v1, target).ExtractMostSignificantBits();
            ulong m2 = Vector128.Equals(v2, target).ExtractMostSignificantBits();
            ulong m3 = Vector128.Equals(v3, target).ExtractMostSignificantBits();
            return m0 | (m1 << 16) | (m2 << 32) | (m3 << 48);
        }

        private void ClassifyUnquoted(ulong delims, ulong crs, ulong lfs)
        {
            ulong structural = delims | crs | lfs;
            ulong lfAfterCr = lfs & ((crs << 1) | _prevCr);
            _ends = structural & ~lfAfterCr;
            _escapes = 0;
            _prevStructural = structural >> 63;
            _prevCr = crs >> 63;
        }

        private void Classify(ulong quotes, ulong delims, ulong crs, ulong lfs, ulong valid)
        {
            ulong inside = PrefixXor(quotes) ^ _insideQuotes;
            ulong opening = quotes & inside;
            ulong closing = quotes & ~inside;
            ulong structural = (delims | crs | lfs) & ~inside;
            ulong afterStructural = (structural << 1) | _prevStructural;
            ulong afterClosing = (closing << 1) | _prevClosing;
            ulong structuralCr = crs & structural;
            ulong lfAfterCr = lfs & ((structuralCr << 1) | _prevCr);

            _ends = structural & ~lfAfterCr;
            _escapes = opening & afterClosing;

            _insideQuotes = (ulong)((long)inside >> 63);
            _prevStructural = structural >> 63;
            _prevClosing = closing >> 63;
            _prevCr = structuralCr >> 63;

            ulong strayOpening = opening & ~(afterStructural | afterClosing);
            ulong textAfterClosing = afterClosing & ~(structural | opening) & valid;
            ulong malformed = strayOpening | textAfterClosing;
            if (malformed == 0)
            {
                return;
            }
            ulong clean = (1UL << BitOperations.TrailingZeroCount(malformed)) - 1;
            _ends &= clean;
            _escapes &= clean;
            _blocked = true;
        }

        // ponytail: CLMUL on x86 only; ARM takes the shift ladder, add Aes.PolynomialMultiplyWideningLower if ARM throughput matters.
        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        private static ulong PrefixXor(ulong bits)
        {
            if (Pclmulqdq.IsSupported)
            {
                return Pclmulqdq.CarrylessMultiply(Vector128.CreateScalar(bits), Vector128<ulong>.AllBitsSet, 0).ToScalar();
            }
            bits ^= bits << 1;
            bits ^= bits << 2;
            bits ^= bits << 4;
            bits ^= bits << 8;
            bits ^= bits << 16;
            bits ^= bits << 32;
            return bits;
        }
    }
}
