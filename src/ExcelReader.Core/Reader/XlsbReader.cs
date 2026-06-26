using System.Diagnostics.CodeAnalysis;

namespace ExcelReader.Core.Reader
{
    // BIFF12 (.xlsb) reader. Shares the same ZIP/OPC container as .xlsx but worksheet parts are
    // binary BIFF12 records instead of XML. Phase 3: enumerator internals. Phase 4 adds IExcelReader,
    // sheet navigation, factory methods, and Excel.Open auto-detect.
    public sealed partial class XlsbReader
    {
        private readonly byte[] _sharedFlat;
        private readonly int[] _sharedOffsets;
        private readonly bool[] _styleIsDate;

        // Phase 3 constructor: accepts pre-parsed components directly (used by tests and Phase 4 factory).
        internal XlsbReader(byte[] sharedFlat, int[] sharedOffsets, bool[] styleIsDate, bool date1904)
        {
            _sharedFlat = sharedFlat;
            _sharedOffsets = sharedOffsets;
            _styleIsDate = styleIsDate;
            IsDate1904 = date1904;
        }

        public bool IsDate1904 { get; }

        internal ReadOnlySpan<byte> SharedSpan => _sharedFlat;

        internal bool IsDateStyle(int style)
        {
            return (uint)style < (uint)_styleIsDate.Length && _styleIsDate[style];
        }


        internal (int Start, int Length) SharedAt(int index)
        {
            if ((uint)index >= (uint)(_sharedOffsets.Length - 1))
            {
                return (0, 0);
            }
            return (_sharedOffsets[index], _sharedOffsets[index + 1] - _sharedOffsets[index]);
        }

        [SuppressMessage("Performance", "HLQ006:GetEnumerator should return a value type",
            Justification = "Public nested enumerator is the standard foreach pattern.")]
        internal Enumerator GetEnumerator(Stream sheetStream)
        {
            return new(this, sheetStream);
        }


        internal Enumerator GetAsyncEnumerator(Stream sheetStream, CancellationToken ct = default)
        {
            return new Enumerator(this, sheetStream, ct);
        }

    }
}
