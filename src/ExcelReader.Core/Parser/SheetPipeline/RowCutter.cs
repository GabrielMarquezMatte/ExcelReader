using ExcelReader.Core.Reader.Xlsb;
using ExcelReader.Core.Reader.Xlsx;

namespace ExcelReader.Core.Parser.SheetPipeline
{
    internal abstract class RowCutter
    {
        // The offset of the last row start past offset 0, or -1 when there is none yet. When `ended` is
        // set, the sheet's rows end at the returned offset and nothing after it is to be read.
        internal abstract int Cut(ReadOnlySpan<byte> data, out bool ended);

        internal virtual void Rebase(int cut)
        {
        }
    }

    internal sealed class XlsxRowCutter(byte[] rowOpen) : RowCutter
    {
        internal override int Cut(ReadOnlySpan<byte> data, out bool ended)
        {
            ended = false;
            ReadOnlySpan<byte> search = data;
            while (true)
            {
                int at = search.LastIndexOf(rowOpen);
                if (at <= 0)
                {
                    return -1;
                }
                int after = at + rowOpen.Length;
                if (after < data.Length && XlsxWorkbook.Enumerator.IsBoundary(data[after]))
                {
                    return at;
                }
                search = search[..at];
            }
        }
    }

    internal sealed class XlsbRowCutter : RowCutter
    {
        private int _walked;
        private int _lastRow = -1;

        internal override int Cut(ReadOnlySpan<byte> data, out bool ended)
        {
            var reader = new Biff12RecordReader(data[_walked..]);
            int before = 0;
            while (reader.TryReadRecord(out int id, out _))
            {
                if (id == Brt.RowHdr)
                {
                    _lastRow = _walked + before;
                }
                else if (id == Brt.EndSheetData)
                {
                    ended = true;
                    return _walked + reader.Position;
                }
                before = reader.Position;
            }
            _walked += reader.Position;
            ended = false;
            return _lastRow > 0 ? _lastRow : -1;
        }

        internal override void Rebase(int cut)
        {
            _walked -= cut;
            _lastRow = -1;
        }
    }
}
