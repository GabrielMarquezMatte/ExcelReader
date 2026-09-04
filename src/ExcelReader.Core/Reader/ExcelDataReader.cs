using System.Data;
using System.Data.Common;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Text;
using ExcelReader.Core.Enums;
using ExcelReader.Core.Parser;
using ExcelReader.Core.ValueObjects;

namespace ExcelReader.Core.Reader
{
    /// <summary>
    /// Adapts the current sheet of an <see cref="IExcelRowReader"/> to <see cref="IDataReader"/>, so it
    /// can feed <c>SqlBulkCopy</c>, <c>DataTable.Load</c>, Dapper, or any other ADO.NET consumer directly.
    /// </summary>
    /// <remarks>
    /// The column shape is fixed once, at construction time, from the header row (or from the first data
    /// row when <c>headerRow</c> is 0) — it never grows as later, wider rows come in. There is no schema
    /// pass, so <see cref="GetFieldType"/>, <see cref="GetValue"/>, and the typed getters all read the
    /// <em>current</em> row's own cell type; a consumer that builds a schema from the first
    /// <see cref="Read"/> (e.g. <c>DataTable.Load</c>) locks in that row's types for the whole load.
    /// <para>
    /// Exposes a single result set: <see cref="NextResult"/> always returns <see langword="false"/>. Use
    /// <see cref="ExcelRowReaderExtensions.Sheets"/> to walk every sheet and construct one
    /// <see cref="ExcelDataReader"/> per sheet instead.
    /// </para>
    /// <para>
    /// Disposing this reader disposes the row enumerator it created, not the <see cref="IExcelRowReader"/>
    /// passed to the constructor — the caller still owns that.
    /// </para>
    /// </remarks>
    public sealed class ExcelDataReader : IDataReader
    {
        private readonly IExcelRowEnumerator _rows;
        private readonly bool _isDate1904;
        private readonly string _sheetName;
        private readonly string?[] _names;
        private readonly Dictionary<string, int> _ordinals;
        private readonly bool _hasPendingRow;
        private bool _pendingConsumed;
        private bool _rowAvailable;
        private bool _disposed;

        /// <summary>
        /// Wraps <paramref name="reader"/>'s current sheet.
        /// </summary>
        /// <param name="reader">The sheet to expose. Not disposed by this reader.</param>
        /// <param name="headerRow">
        /// The 1-based row holding column names. Pass 0 for a header-less sheet, whose columns come back
        /// named <c>"Column0"</c>, <c>"Column1"</c>, ... and whose count is taken from the first data row.
        /// </param>
        /// <exception cref="ArgumentNullException"><paramref name="reader"/> is null.</exception>
        /// <exception cref="ArgumentOutOfRangeException"><paramref name="headerRow"/> is negative.</exception>
        public ExcelDataReader(IExcelRowReader reader, int headerRow = 1)
        {
            ArgumentNullException.ThrowIfNull(reader);
            ArgumentOutOfRangeException.ThrowIfNegative(headerRow);
            _isDate1904 = reader.IsDate1904;
            _sheetName = reader.SheetName;
            _rows = reader.GetEnumerator();
            _pendingConsumed = true;
            if (headerRow > 0 && SchemaInference.TrySkipToHeaderRow(_rows, headerRow, out _))
            {
                _names = ReadHeaderNames(_rows.Current);
            }
            else if (headerRow > 0)
            {
                // Sheet has fewer rows than headerRow: an empty, columnless result set.
                _names = [];
            }
            else
            {
                _hasPendingRow = _rows.MoveNext();
                _pendingConsumed = false;
                _names = _hasPendingRow ? new string?[_rows.Current.ColumnCount] : [];
            }
            _ordinals = BuildOrdinals(_names);
        }

        private static string?[] ReadHeaderNames(Row header)
        {
            string?[] names = new string?[header.ColumnCount];
            foreach (RowCell cell in header.Cells)
            {
                string text = cell.Value.GetString().Trim();
                names[cell.ColumnIndex] = text.Length == 0 ? null : text;
            }
            return names;
        }

        private static Dictionary<string, int> BuildOrdinals(string?[] names)
        {
            var map = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            for (int i = 0; i < names.Length; i++)
            {
                if (names[i] is { } name)
                {
                    map[name] = i; // a repeated header name keeps its last column
                }
            }
            return map;
        }

        /// <inheritdoc/>
        public int FieldCount => _names.Length;

        /// <inheritdoc/>
        public int Depth => 0;

        /// <inheritdoc/>
        public bool IsClosed => _disposed;

        /// <inheritdoc/>
        public int RecordsAffected => -1;

        /// <inheritdoc/>
        public object this[int i] => GetValue(i);

        /// <inheritdoc/>
        public object this[string name] => GetValue(GetOrdinal(name));

        /// <inheritdoc/>
        public bool Read()
        {
            if (!_pendingConsumed)
            {
                _pendingConsumed = true;
                _rowAvailable = _hasPendingRow;
                return _rowAvailable;
            }
            _rowAvailable = _rows.MoveNext();
            return _rowAvailable;
        }

        /// <summary>Always returns <see langword="false"/> — see the type-level remarks.</summary>
        public bool NextResult()
        {
            return false;
        }

        /// <inheritdoc/>
        public void Close()
        {
            Dispose();
        }

        /// <inheritdoc/>
        public void Dispose()
        {
            if (_disposed)
            {
                return;
            }
            _disposed = true;
            _rows.Dispose();
        }

        /// <inheritdoc/>
        public string GetName(int i)
        {
            return _names[i] ?? $"Column{i}";
        }

        /// <inheritdoc/>
        public int GetOrdinal(string name)
        {
            return _ordinals.TryGetValue(name, out int i) ? i : throw new KeyNotFoundException($"Column '{name}' was not found.");
        }

        /// <inheritdoc/>
        public string GetDataTypeName(int i)
        {
            return GetFieldType(i).Name;
        }

        /// <inheritdoc/>
        // IDataRecord.GetFieldType's return value carries this same annotation in the BCL (DataTable's
        // schema machinery inspects a column's Type via its public fields/properties). An override must
        // repeat an interface member's DynamicallyAccessedMembersAttribute exactly — IL2093 otherwise —
        // even though every branch below returns a closed, well-known type that needs no such access.
        [return: DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicFields | DynamicallyAccessedMemberTypes.PublicProperties)]
        public Type GetFieldType(int i)
        {
            return CurrentCell(i).Type switch
            {
                CellType.Number => typeof(double),
                CellType.Date => typeof(DateTime),
                CellType.Boolean => typeof(bool),
                _ => typeof(string),
            };
        }

        /// <inheritdoc/>
        public object GetValue(int i)
        {
            Cell cell = CurrentCell(i);
            return cell.Type switch
            {
                CellType.Empty => DBNull.Value,
                CellType.Number => cell.TryGetDouble(out double d) ? d : cell.GetString(),
                CellType.Date => cell.TryGetDateTime(_isDate1904, out DateTime dt) ? dt : cell.GetString(),
                CellType.Boolean => ExcelCellReaders.Bool(in cell, _isDate1904, CultureInfo.InvariantCulture, out bool b) ? b : cell.GetString(),
                _ => cell.GetString(),
            };
        }

        /// <inheritdoc/>
        public int GetValues(object[] values)
        {
            ArgumentNullException.ThrowIfNull(values);
            int n = Math.Min(values.Length, FieldCount);
            for (int i = 0; i < n; i++)
            {
                values[i] = GetValue(i);
            }
            return n;
        }

        /// <inheritdoc/>
        public bool IsDBNull(int i)
        {
            return CurrentCell(i).Type == CellType.Empty;
        }

        /// <inheritdoc/>
        public string GetString(int i)
        {
            return CurrentCell(i).GetString();
        }

        /// <inheritdoc/>
        public bool GetBoolean(int i)
        {
            Cell cell = CurrentCell(i);
            return ExcelCellReaders.Bool(in cell, _isDate1904, CultureInfo.InvariantCulture, out bool value)
                ? value
                : throw new FormatException($"Column {i} is not a valid boolean.");
        }

        /// <inheritdoc/>
        public DateTime GetDateTime(int i)
        {
            Cell cell = CurrentCell(i);
            return cell.TryGetDateTime(_isDate1904, out DateTime value)
                ? value
                : throw new FormatException($"Column {i} is not a valid date/time.");
        }

        /// <inheritdoc/>
        public byte GetByte(int i)
        {
            return GetParsed<byte>(i);
        }

        /// <inheritdoc/>
        public short GetInt16(int i)
        {
            return GetParsed<short>(i);
        }

        /// <inheritdoc/>
        public int GetInt32(int i)
        {
            return GetParsed<int>(i);
        }

        /// <inheritdoc/>
        public long GetInt64(int i)
        {
            return GetParsed<long>(i);
        }

        /// <inheritdoc/>
        public float GetFloat(int i)
        {
            return GetParsed<float>(i);
        }

        /// <inheritdoc/>
        public double GetDouble(int i)
        {
            return GetParsed<double>(i);
        }

        /// <inheritdoc/>
        public decimal GetDecimal(int i)
        {
            return GetParsed<decimal>(i);
        }

        /// <inheritdoc/>
        public Guid GetGuid(int i)
        {
            Cell cell = CurrentCell(i);
#if NET8_0
            return ExcelCellReaders.Guid(in cell, _isDate1904, CultureInfo.InvariantCulture, out Guid value)
#else
            return cell.TryParse(CultureInfo.InvariantCulture, out Guid value)
#endif
                ? value
                : throw new FormatException($"Column {i} is not a valid Guid.");
        }

        /// <inheritdoc/>
        public char GetChar(int i)
        {
            string s = GetString(i);
            return s.Length > 0 ? s[0] : throw new InvalidCastException($"Column {i} is empty; no char value.");
        }

        /// <inheritdoc/>
        public long GetBytes(int i, long fieldOffset, byte[]? buffer, int bufferoffset, int length)
        {
            byte[] bytes = IsDBNull(i) ? [] : Encoding.UTF8.GetBytes(GetString(i));
            return CopySlice(bytes.Length, fieldOffset, buffer?.Length, length,
                toCopy => Array.Copy(bytes, (int)fieldOffset, buffer!, bufferoffset, toCopy));
        }

        /// <inheritdoc/>
        public long GetChars(int i, long fieldoffset, char[]? buffer, int bufferoffset, int length)
        {
            string s = GetString(i);
            return CopySlice(s.Length, fieldoffset, buffer?.Length, length,
                toCopy => s.CopyTo((int)fieldoffset, buffer!, bufferoffset, toCopy));
        }

        private static long CopySlice(int sourceLength, long fieldOffset, int? bufferLength, int length, Action<int> copy)
        {
            if (bufferLength is null)
            {
                return sourceLength;
            }
            int available = (int)Math.Max(0, sourceLength - fieldOffset);
            int toCopy = Math.Max(0, Math.Min(length, available));
            if (toCopy > 0)
            {
                copy(toCopy);
            }
            return toCopy;
        }

        /// <inheritdoc/>
        public IDataReader GetData(int i)
        {
            throw new NotSupportedException("ExcelDataReader has no nested result sets.");
        }

        /// <inheritdoc/>
        // IL2111: DataColumnCollection.Add(string, Type)'s `type` parameter carries the same
        // PublicFields|PublicProperties DynamicallyAccessedMembersAttribute as GetFieldType's return
        // above. Passing `typeof(Type)` itself as that argument — the "DataType" schema column's own
        // type is System.Type, describing the DataType column, not a real cell type — trips a known
        // ILC/trimmer quirk: satisfying that annotation for the argument `Type` requires inspecting
        // Type's own public properties, one of which (TypeInitializer) is itself DAM-annotated, and the
        // linker can't statically prove that recursive requirement holds. Safe here: this DataTable is
        // schema metadata for DataTable.Load/FillSchema; nothing ever reflects over the value stored in
        // this column via the annotated members.
        [UnconditionalSuppressMessage("Trimming", "IL2111",
            Justification = "typeof(Type) as the 'DataType' schema column's own type is a metadata literal, never reflected over.")]
        public DataTable GetSchemaTable()
        {
            // DataTable.Load (via DbDataAdapter.FillSchema) reads this exact standard shape
            // (System.Data.Common.SchemaTableColumn/SchemaTableOptionalColumn) — trimming it down to
            // just the columns this reader actually varies (ColumnName/ColumnOrdinal/DataType/AllowDBNull)
            // makes DataTable.Load throw internally, since it indexes the rest unconditionally.
            var table = new DataTable();
            table.Columns.Add(SchemaTableColumn.ColumnName, typeof(string));
            table.Columns.Add(SchemaTableColumn.ColumnOrdinal, typeof(int));
            table.Columns.Add(SchemaTableColumn.ColumnSize, typeof(int));
            table.Columns.Add(SchemaTableColumn.NumericPrecision, typeof(short));
            table.Columns.Add(SchemaTableColumn.NumericScale, typeof(short));
            table.Columns.Add(SchemaTableColumn.DataType, typeof(Type));
            table.Columns.Add(SchemaTableColumn.IsLong, typeof(bool));
            table.Columns.Add(SchemaTableColumn.AllowDBNull, typeof(bool));
            table.Columns.Add(SchemaTableColumn.IsUnique, typeof(bool));
            table.Columns.Add(SchemaTableColumn.IsKey, typeof(bool));
            table.Columns.Add(SchemaTableOptionalColumn.IsAutoIncrement, typeof(bool));
            table.Columns.Add(SchemaTableOptionalColumn.IsReadOnly, typeof(bool));
            table.Columns.Add(SchemaTableOptionalColumn.IsRowVersion, typeof(bool));
            table.Columns.Add(SchemaTableOptionalColumn.IsHidden, typeof(bool));
            table.Columns.Add(SchemaTableColumn.IsExpression, typeof(bool));
            table.Columns.Add(SchemaTableColumn.IsAliased, typeof(bool));
            table.Columns.Add(SchemaTableColumn.BaseSchemaName, typeof(string));
            table.Columns.Add(SchemaTableOptionalColumn.BaseCatalogName, typeof(string));
            table.Columns.Add(SchemaTableColumn.BaseTableName, typeof(string));
            table.Columns.Add(SchemaTableColumn.BaseColumnName, typeof(string));
            for (int i = 0; i < FieldCount; i++)
            {
                // No type is known before the first row is read; guess string, the type every cell can
                // represent verbatim (same fallback SchemaInference uses for an unresolved column).
                Type type = _rowAvailable ? GetFieldType(i) : typeof(string);
                string name = GetName(i);
                DataRow row = table.NewRow();
                row[SchemaTableColumn.ColumnName] = name;
                row[SchemaTableColumn.ColumnOrdinal] = i;
                row[SchemaTableColumn.ColumnSize] = -1;
                row[SchemaTableColumn.NumericPrecision] = (short)0;
                row[SchemaTableColumn.NumericScale] = (short)0;
                row[SchemaTableColumn.DataType] = type;
                row[SchemaTableColumn.IsLong] = false;
                row[SchemaTableColumn.AllowDBNull] = true;
                row[SchemaTableColumn.IsUnique] = false;
                row[SchemaTableColumn.IsKey] = false;
                row[SchemaTableOptionalColumn.IsAutoIncrement] = false;
                row[SchemaTableOptionalColumn.IsReadOnly] = true;
                row[SchemaTableOptionalColumn.IsRowVersion] = false;
                row[SchemaTableOptionalColumn.IsHidden] = false;
                row[SchemaTableColumn.IsExpression] = false;
                row[SchemaTableColumn.IsAliased] = false;
                row[SchemaTableColumn.BaseSchemaName] = DBNull.Value;
                row[SchemaTableOptionalColumn.BaseCatalogName] = DBNull.Value;
                row[SchemaTableColumn.BaseTableName] = _sheetName;
                row[SchemaTableColumn.BaseColumnName] = name;
                table.Rows.Add(row);
            }
            return table;
        }

        private Cell CurrentCell(int i)
        {
            ArgumentOutOfRangeException.ThrowIfNegative(i);
            ArgumentOutOfRangeException.ThrowIfGreaterThanOrEqual(i, FieldCount);
            if (!_rowAvailable)
            {
                throw new InvalidOperationException("No current row; call Read() first.");
            }
            return _rows.Current[i];
        }

        private T GetParsed<T>(int i) where T : struct, IUtf8SpanParsable<T>
        {
            Cell cell = CurrentCell(i);
            return cell.TryParse(CultureInfo.InvariantCulture, out T value)
                ? value
                : throw new FormatException($"Column {i} is not a valid {typeof(T).Name}.");
        }
    }
}
