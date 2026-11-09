using System.Runtime.ExceptionServices;
using ExcelReader.Core.Reader.Internal;

namespace ExcelReader.Core.Reader.Columnar
{
    /// <summary>Appends rows to one builder per column; a parallel parse keeps one table per partition.</summary>
    internal sealed class TypedTableBuilder : IDisposable
    {
        private readonly TypedColumnBuilder[] _columns;

        internal TypedTableBuilder(TypedColumnSpec[] specs)
        {
            _columns = new TypedColumnBuilder[specs.Length];
            for (int i = 0; i < specs.Length; i++)
            {
                _columns[i] = new TypedColumnBuilder(specs[i]);
            }
        }

        internal long RowCount { get; private set; }

        internal bool TryAppend(in Row row, int[] columnIndices, bool isDate1904, out int failedColumn)
        {
            ReadOnlySpan<CellDesc> cells = row.Descriptors;
            for (int i = 0; i < _columns.Length; i++)
            {
                int index = row.IndexOf(columnIndices[i]);
                if (index >= 0 && _columns[i].CanAppendDirect(in cells[index]) && _columns[i].TryAppendDirect(in cells[index], in row))
                {
                    continue;
                }
                if (!_columns[i].AppendFrom(CellAt(in row, index), isDate1904))
                {
                    failedColumn = i;
                    return false;
                }
            }
            RowCount++;
            failedColumn = -1;
            return true;
        }

        private static Cell CellAt(in Row row, int index)
        {
            if (index < 0)
            {
                return new Cell(CellType.Empty, default);
            }
            return row.CellAt(index);
        }

        internal ColumnarTable Build()
        {
            return Build([this], degreeOfParallelism: 1);
        }

        internal static ColumnarTable Build(IReadOnlyList<TypedTableBuilder> parts, int degreeOfParallelism)
        {
            long rowCount = 0;
            foreach (TypedTableBuilder part in parts)
            {
                rowCount += part.RowCount;
            }
            int columnCount = parts[0]._columns.Length;
            ColumnarTable table = new(rowCount, new BuiltColumn[columnCount]);
            try
            {
                if (degreeOfParallelism == 1)
                {
                    for (int c = 0; c < columnCount; c++)
                    {
                        table.Columns[c] = BuildColumn(parts, c);
                    }
                    return table;
                }
                Parallel.For(0, columnCount, new ParallelOptions { MaxDegreeOfParallelism = degreeOfParallelism },
                    c => table.Columns[c] = BuildColumn(parts, c));
                return table;
            }
            catch (AggregateException exception) when (exception.InnerExceptions.Count == 1)
            {
                table.Dispose();
                ExceptionDispatchInfo.Throw(exception.InnerExceptions[0]);
                throw;
            }
            catch
            {
                table.Dispose();
                throw;
            }
        }

        private static BuiltColumn BuildColumn(IReadOnlyList<TypedTableBuilder> parts, int index)
        {
            TypedColumnBuilder[] column = new TypedColumnBuilder[parts.Count];
            for (int p = 0; p < parts.Count; p++)
            {
                column[p] = parts[p]._columns[index];
            }
            return TypedColumnBuilder.Build(column);
        }

        public void Dispose()
        {
            foreach (TypedColumnBuilder column in _columns)
            {
                column.Dispose();
            }
        }
    }
}
