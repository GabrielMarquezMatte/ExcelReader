using System.Diagnostics.CodeAnalysis;
using ExcelReader.NetFramework.Interop;
using ExcelReader.NetFramework.Mapping;

namespace ExcelReader.NetFramework
{
    /// <summary>A sheet or table of an open workbook.</summary>
    public readonly struct ExcelSheet : IEquatable<ExcelSheet>
    {
        internal ExcelSheet(ExcelWorkbook workbook, int id)
        {
            Workbook = workbook;
            Id = id;
        }

        /// <summary>The workbook this sheet belongs to.</summary>
        public ExcelWorkbook Workbook { get; }

        internal int Id { get; }

        /// <summary>Lazily enumerates the raw rows of the sheet or table; each enumeration opens its own cursor.</summary>
        /// <returns>The rows, header included.</returns>
        public IEnumerable<ExcelRow> ReadRows()
        {
            using RowsHandle rows = RowsHandle.Open(Workbook, Id);
            while (true)
            {
                ExcelRow? row = rows.Next();
                if (row is null)
                {
                    yield break;
                }

                yield return row;
            }
        }

        /// <summary>
        /// Reads every data row into a new <typeparamref name="T"/>, matching public settable properties to header
        /// columns by <see cref="ExcelColumnAttribute"/> names or the property name. Properties whose column is
        /// absent, and cells that are empty or do not convert, keep the default value.
        /// </summary>
        /// <typeparam name="T">A class with a public parameterless constructor.</typeparam>
        /// <param name="headerRow">The 1-based row holding the column headers; data starts on the next row.</param>
        /// <param name="degreeOfParallelism">The number of threads to parse with; 0 uses the processor count.</param>
        /// <returns>One item per data row, in sheet order.</returns>
        /// <exception cref="NotSupportedException">A property of <typeparamref name="T"/> has an unsupported type.</exception>
        /// <exception cref="ExcelReaderException">A required column or value is missing, or a value does not fit its property.</exception>
        [SuppressMessage("Design", "CA1002", Justification = "The caller owns the result; List<T> is the documented return type.")]
        [SuppressMessage("Design", "MA0016", Justification = "The caller owns the result; List<T> is the documented return type.")]
        public unsafe List<T> Read<T>(int headerRow = 1, int degreeOfParallelism = 0)
            where T : new()
        {
            if (headerRow < 1)
            {
                throw new ArgumentOutOfRangeException(nameof(headerRow), headerRow, "Typed reads match columns by name and need a header row of 1 or more.");
            }

            if (degreeOfParallelism < 0)
            {
                throw new ArgumentOutOfRangeException(nameof(degreeOfParallelism), degreeOfParallelism, "Use 0 for the processor count or a positive thread count.");
            }

            TypeMap<T> map = TypeMap<T>.Instance;
            Dictionary<string, int> header = HeaderReader.Read(this, headerRow);
            using NativeSpecs specs = NativeSpecs.Create(map, header, out PropertyBinder<T>[] bound);
            if (bound.Length == 0)
            {
                throw new ExcelReaderException(ExcelStatus.InvalidArgument, $"None of the columns of '{typeof(T).Name}' is in header row {headerRow}.");
            }

            WorkbookHandle handle = Workbook.Handle;
            bool added = false;
            XlTable table = default;
            try
            {
                handle.DangerousAddRef(ref added);
                NativeError.ThrowIfFailed(NativeMethods.xl_parse_typed(handle.DangerousGetHandle(), Id, specs.Pointer, specs.Count, headerRow, degreeOfParallelism, &table));
                T[] items = new T[checked((int)table.RowCount)];
                for (int i = 0; i < items.Length; i++)
                {
                    items[i] = new T();
                }

                TableFiller.Fill(bound, in table, items, 0);
                return new List<T>(items);
            }
            finally
            {
                NativeMethods.xl_free_table(&table);
                if (added)
                {
                    handle.DangerousRelease();
                }
            }
        }

        /// <summary>Compares two sheets for equality.</summary>
        /// <param name="left">The first sheet.</param>
        /// <param name="right">The second sheet.</param>
        /// <returns>True when both refer to the same sheet of the same workbook.</returns>
        public static bool operator ==(ExcelSheet left, ExcelSheet right)
        {
            return left.Equals(right);
        }

        /// <summary>Compares two sheets for inequality.</summary>
        /// <param name="left">The first sheet.</param>
        /// <param name="right">The second sheet.</param>
        /// <returns>True when the sheets differ.</returns>
        public static bool operator !=(ExcelSheet left, ExcelSheet right)
        {
            return !left.Equals(right);
        }

        /// <inheritdoc />
        public bool Equals(ExcelSheet other)
        {
            return ReferenceEquals(Workbook, other.Workbook) && Id == other.Id;
        }

        /// <inheritdoc />
        public override bool Equals(object? obj)
        {
            return obj is ExcelSheet other && Equals(other);
        }

        /// <inheritdoc />
        public override int GetHashCode()
        {
            return Id;
        }
    }
}
