namespace ExcelReader.Core.Reader.Internal
{
    internal sealed class TableRowEnumerator : IExcelRowEnumerator
    {
        private readonly IExcelRowEnumerator _rows;
        private readonly IRowIndexedEnumerator _indexed;
        private readonly int _firstRow;
        private readonly int _lastDataRow;
        private readonly int _firstColumn;
        private readonly int _lastColumn;
        private int _previousRow = -1;
        private bool _done;

        internal TableRowEnumerator(IExcelRowEnumerator rows, int firstRow, int lastDataRow, int firstColumn, int lastColumn)
        {
            _rows = rows;
            _indexed = (IRowIndexedEnumerator)rows;
            _indexed.EnableRowIndex();
            _firstRow = firstRow;
            _lastDataRow = lastDataRow;
            _firstColumn = firstColumn;
            _lastColumn = lastColumn;
        }

        public Row Current
        {
            get
            {
                return _rows.Current.Slice(_firstColumn, _lastColumn);
            }
        }

        public bool MoveNext()
        {
            while (!_done && _rows.MoveNext())
            {
                if (IsInTable())
                {
                    return true;
                }
            }
            _done = true;
            return false;
        }

        public async ValueTask<bool> MoveNextAsync()
        {
            while (!_done && await _rows.MoveNextAsync().ConfigureAwait(false))
            {
                if (IsInTable())
                {
                    return true;
                }
            }
            _done = true;
            return false;
        }

        public void Dispose()
        {
            _rows.Dispose();
        }

        public ValueTask DisposeAsync()
        {
            return _rows.DisposeAsync();
        }

        private bool IsInTable()
        {
            int row = _indexed.RowIndex;
            if (row < 0 || row >= ExcelLimits.MaxRows || row <= _previousRow)
            {
                throw new InvalidDataException($"Row index {row} is out of range or not after the previous row {_previousRow}.");
            }
            _previousRow = row;
            if (row > _lastDataRow)
            {
                _done = true;
                return false;
            }
            return row >= _firstRow;
        }
    }
}
