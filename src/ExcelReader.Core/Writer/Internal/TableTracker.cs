using ExcelReader.Core.Reader.Internal;

namespace ExcelReader.Core.Writer.Internal
{
    internal sealed class TableTracker
    {
        private readonly List<WrittenTable> _finished = [];
        private WrittenTable? _open;

        internal IReadOnlyList<WrittenTable> Finished
        {
            get
            {
                return _finished;
            }
        }

        internal void RequireNoneOpen()
        {
            if (_open is not null)
            {
                throw new InvalidOperationException($"Table '{_open.Name}' is still open on this sheet; call EndTable first.");
            }
        }

        internal void Begin(int id, string name, string[] columns, ExcelTableOptions options, int headerRow)
        {
            RequireNoneOpen();
            _open = new WrittenTable(id, name, columns, options, headerRow, headerRow);
        }

        internal void End(int lastWrittenRow)
        {
            if (_open is null)
            {
                throw new InvalidOperationException("No table is open on this sheet.");
            }
            CloseOpen(lastWrittenRow);
        }

        internal void CloseOpen(int lastWrittenRow)
        {
            if (_open is null)
            {
                return;
            }
            _finished.Add(_open with { LastRow = Math.Max(lastWrittenRow, _open.HeaderRow + 1) });
            _open = null;
        }

        internal static void RequireRoom(int nextRowIndex)
        {
            if (nextRowIndex > ExcelLimits.MaxRows - 2)
            {
                ExcelLimits.ThrowRowLimit(nextRowIndex + 2L);
            }
        }

        internal static void WriteHeader<TRow>(TRow row, int firstColumn, string[] columns)
            where TRow : IRowWriter
        {
            if (firstColumn > 0)
            {
                row.Skip(firstColumn);
            }
            foreach (string column in columns)
            {
                row.Write(column);
            }
        }
    }
}
