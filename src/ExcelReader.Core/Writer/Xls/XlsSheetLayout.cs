using System.Runtime.InteropServices;
using ExcelReader.Core.Writer.Internal;

namespace ExcelReader.Core.Writer.Xls
{
    /// <summary>
    /// Lays out one worksheet substream in the order [MS-XLS] requires: INDEX, the sheet settings,
    /// DEFCOLWIDTH and COLINFO, DIMENSIONS, the cell table in blocks of up to 32 rows (their ROW records,
    /// their cells, then a DBCELL pointing back at both), WINDOW2 and EOF. The cells stay in the sheet's
    /// own buffer and are written between the framing records rather than copied.
    /// </summary>
    internal sealed class XlsSheetLayout : IDisposable
    {
        private const int RowsPerBlock = 32;
        private const int RowRecordLength = 4 + 16;

        private readonly BiffBuffer _frame = new(1024);
        private readonly ReadOnlyMemory<byte> _cells;
        private readonly List<Segment> _segments = [];
        private readonly List<(int Position, int Relative)> _filePointers = [];

        internal XlsSheetLayout(XlsSheetWriter sheet, bool selected)
        {
            _cells = sheet.CellsMemory;
            ReadOnlySpan<XlsRowExtent> rows = sheet.Rows;
            int blockCount = (rows.Length + RowsPerBlock - 1) / RowsPerBlock;

            BiffRecordWriter.WriteBof(_frame, BiffRecord.SubstreamWorksheet);
            int index = _frame.BeginRecord(BiffRecord.Index);
            _frame.WriteI32(0);
            _frame.WriteI32(0);
            _frame.WriteI32(rows.Length);
            int defColWidthPointer = _frame.Length;
            _frame.WriteI32(0);
            int dbCellPointers = _frame.Length;
            for (int i = 0; i < blockCount; i++)
            {
                _frame.WriteI32(0);
            }
            _frame.EndRecord(index);

            BiffRecordWriter.WriteSheetSettings(_frame);
            _filePointers.Add((defColWidthPointer, _frame.Length));
            BiffRecordWriter.WriteDefColWidth(_frame);
            _frame.Write(sheet.ColInfoMemory.Span);
            BiffRecordWriter.WriteDimension(_frame, sheet.RowCount, sheet.ColCount);
            AddFrame(0);

            int position = _frame.Length;
            for (int block = 0; block < blockCount; block++)
            {
                int first = block * RowsPerBlock;
                int count = Math.Min(RowsPerBlock, rows.Length - first);
                int cellsStart = rows[first].CellStart;
                int cellsEnd = first + count < rows.Length ? rows[first + count].CellStart : _cells.Length;

                int rowsFrame = _frame.Length;
                for (int i = first; i < first + count; i++)
                {
                    BiffRecordWriter.WriteRow(_frame, i, rows[i].ColMic, rows[i].ColMac);
                }
                AddFrame(rowsFrame);
                _segments.Add(new Segment(Cells: true, cellsStart, cellsEnd - cellsStart));
                int blockLength = (count * RowRecordLength) + (cellsEnd - cellsStart);
                position += blockLength;

                _filePointers.Add((dbCellPointers + (4 * block), position));
                int dbCellFrame = _frame.Length;
                int dbCell = _frame.BeginRecord(BiffRecord.DbCell);
                _frame.WriteI32(blockLength);
                // The first offset runs from the block's second ROW record to the first row's cells;
                // each later one is the length of the previous row's cells.
                _frame.WriteU16((count - 1) * RowRecordLength);
                for (int i = first + 1; i < first + count; i++)
                {
                    _frame.WriteU16(Math.Min(rows[i].CellStart - rows[i - 1].CellStart, ushort.MaxValue));
                }
                _frame.EndRecord(dbCell);
                AddFrame(dbCellFrame);
                position += _frame.Length - dbCellFrame;
            }

            int tailFrame = _frame.Length;
            BiffRecordWriter.WriteWindow2(_frame, selected);
            BiffRecordWriter.WriteEof(_frame);
            AddFrame(tailFrame);
            Length = position + (_frame.Length - tailFrame);
        }

        internal int Length { get; }

        /// <summary>INDEX holds absolute stream positions, known only once the globals have been sized.</summary>
        internal void Place(int substreamStart)
        {
            foreach ((int pointer, int relative) in _filePointers)
            {
                _frame.PatchI32(pointer, substreamStart + relative);
            }
        }

        internal void WriteTo(Stream destination)
        {
            foreach (Segment segment in _segments)
            {
                destination.Write(Slice(segment).Span);
            }
        }

        internal async ValueTask WriteToAsync(Stream destination, CancellationToken ct)
        {
            foreach (Segment segment in _segments)
            {
                await destination.WriteAsync(Slice(segment), ct).ConfigureAwait(false);
            }
        }

        public void Dispose()
        {
            _frame.Dispose();
        }

        private void AddFrame(int start)
        {
            _segments.Add(new Segment(Cells: false, start, _frame.Length - start));
        }

        private ReadOnlyMemory<byte> Slice(Segment segment)
        {
            return (segment.Cells ? _cells : _frame.Memory).Slice(segment.Start, segment.Length);
        }

        [StructLayout(LayoutKind.Auto)]
        private readonly record struct Segment(bool Cells, int Start, int Length);
    }
}
