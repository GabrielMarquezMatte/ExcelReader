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

        /// <summary>
        /// Lazily reads the data rows into batches of at most <paramref name="batchRows"/> items of <typeparamref name="T"/>,
        /// with the same column matching as <see cref="Read{T}"/>. The native reader opens on the first batch and closes
        /// when the enumeration ends or is disposed.
        /// </summary>
        /// <typeparam name="T">A class with a public parameterless constructor.</typeparam>
        /// <param name="batchRows">The maximum number of rows per batch; at least 1.</param>
        /// <param name="headerRow">The 1-based row holding the column headers; data starts on the next row.</param>
        /// <returns>The batches, in sheet order.</returns>
        /// <exception cref="NotSupportedException">A property of <typeparamref name="T"/> has an unsupported type.</exception>
        /// <exception cref="ExcelReaderException">A required column or value is missing, or a value does not fit its property.</exception>
        [SuppressMessage("Design", "CA1002", Justification = "The caller owns each batch; List<T> is the documented element type.")]
        [SuppressMessage("Design", "MA0016", Justification = "The caller owns each batch; List<T> is the documented element type.")]
        public IEnumerable<List<T>> ReadBatches<T>(int batchRows, int headerRow = 1)
            where T : new()
        {
            if (batchRows < 1)
            {
                throw new ArgumentOutOfRangeException(nameof(batchRows), batchRows, "A batch needs at least one row.");
            }

            if (headerRow < 1)
            {
                throw new ArgumentOutOfRangeException(nameof(headerRow), headerRow, "Typed reads match columns by name and need a header row of 1 or more.");
            }

            return ReadBatchesCore<T>(batchRows, headerRow);
        }

        /// <summary>Infers the type and nullability of each column from the first rows of the sheet.</summary>
        /// <param name="headerRow">The 1-based row holding the column headers; 0 when the sheet has none.</param>
        /// <param name="sampleRows">The number of data rows to sample; at least 1.</param>
        /// <param name="parseText">True to also infer numbers, booleans and dates from text cells.</param>
        /// <returns>One entry per column, in column order.</returns>
        public unsafe IReadOnlyList<ExcelColumnSchema> InferSchema(int headerRow = 1, int sampleRows = 1000, bool parseText = false)
        {
            if (headerRow < 0)
            {
                throw new ArgumentOutOfRangeException(nameof(headerRow), headerRow, "Use 0 for no header row or the 1-based header row.");
            }

            if (sampleRows < 1)
            {
                throw new ArgumentOutOfRangeException(nameof(sampleRows), sampleRows, "Sample at least one row.");
            }

            WorkbookHandle handle = Workbook.Handle;
            bool added = false;
            XlInferredSchema schema = default;
            try
            {
                handle.DangerousAddRef(ref added);
                int flags = parseText ? Xl.XL_INFER_PARSE_TEXT : 0;
                NativeError.ThrowIfFailed(NativeMethods.xl_infer_schema(handle.DangerousGetHandle(), Id, headerRow, sampleRows, flags, &schema));
                List<ExcelColumnSchema> columns = new List<ExcelColumnSchema>(schema.ColumnCount);
                for (int i = 0; i < schema.ColumnCount; i++)
                {
                    XlColumnSpec spec = schema.Columns[i];
                    string? name = spec.NameCount > 0 ? Utf8.Decode(spec.Names[0], spec.NameLens[0]) : null;
                    columns.Add(new ExcelColumnSchema(name, (ExcelColumnType)spec.Type, spec.Nullable != 0));
                }

                return columns;
            }
            finally
            {
                NativeMethods.xl_free_schema(&schema);
                if (added)
                {
                    handle.DangerousRelease();
                }
            }
        }

        private IEnumerable<List<T>> ReadBatchesCore<T>(int batchRows, int headerRow)
            where T : new()
        {
            TypeMap<T> map = TypeMap<T>.Instance;
            Dictionary<string, int> header = HeaderReader.Read(this, headerRow);
            using NativeSpecs specs = NativeSpecs.Create(map, header, out PropertyBinder<T>[] bound);
            if (bound.Length == 0)
            {
                throw new ExcelReaderException(ExcelStatus.InvalidArgument, $"None of the columns of '{typeof(T).Name}' is in header row {headerRow}.");
            }

            using TypedReaderHandle reader = TypedReaderHandle.Open(Workbook, Id, specs, headerRow, batchRows);
            while (true)
            {
                List<T>? batch = reader.Next(bound);
                if (batch is null)
                {
                    yield break;
                }

                yield return batch;
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
