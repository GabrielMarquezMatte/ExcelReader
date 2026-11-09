using System.Runtime.CompilerServices;

namespace ExcelReader.Core.Reader.Internal
{
    /// <summary>
    /// One <see cref="Utf8StringCache"/> per column: a mostly-unique column neither evicts a categorical
    /// column's values nor keeps paying for lookups once its own cache pauses. Not thread-safe; one per enumerator.
    /// </summary>
    internal sealed class ColumnStringCaches
    {
        private Utf8StringCache?[] _columns = new Utf8StringCache?[16];

        /// <summary>The column's cache, or null while it is paused.</summary>
        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        internal Utf8StringCache? ForColumn(int column)
        {
            Utf8StringCache?[] columns = _columns;
            if ((uint)column >= (uint)columns.Length)
            {
                return ForNewColumn(column);
            }
            Utf8StringCache cache = columns[column] ??= new Utf8StringCache();
            return cache.TryUse() ? cache : null;
        }

        [MethodImpl(MethodImplOptions.NoInlining)]
        private Utf8StringCache? ForNewColumn(int column)
        {
            if ((uint)column >= ExcelLimits.MaxColumns)
            {
                return null;
            }
            Array.Resize(ref _columns, Math.Min(ExcelLimits.MaxColumns, Math.Max(column + 1, _columns.Length * 2)));
            return ForColumn(column);
        }
    }
}
