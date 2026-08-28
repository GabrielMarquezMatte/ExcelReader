using System.Diagnostics.CodeAnalysis;
using System.IO.Compression;
using System.Linq.Expressions;
using System.Reflection;
using ExcelReader.Core.Parser;

namespace ExcelReader.Core.Writer
{
    /// <summary>
    /// Writes plain-old-CLR-object records to a workbook as sheets: each call to <c>WriteSheetAsync</c>
    /// writes a new sheet, with a header row (from each property's <c>[ExcelColumn]</c> name or, failing
    /// that, the property name) followed by one row per record, one column per public readable property
    /// of the record type. Each call targets a new sheet, so one workbook can hold sheets of different
    /// record types. Works uniformly across XLSX, XLSB, XLS and CSV via the
    /// <see cref="IWorkbookWriter{TSheet}"/>/<see cref="ISheetWriter{TRow}"/>/<see cref="IRowWriter"/>
    /// abstractions. Prefer the <see cref="RecordWriter"/> factory methods over constructing this type
    /// directly. Headers written here match the property names/aliases that <c>ExcelParser&lt;T&gt;</c>
    /// looks for, so column order does not need to match between writing and reading.
    /// </summary>
    /// <typeparam name="TSheet">The concrete sheet writer type.</typeparam>
    /// <typeparam name="TRow">The concrete row writer type.</typeparam>
    public sealed class WorkbookRecordWriter<TSheet, TRow> : IAsyncDisposable where TSheet : ISheetWriter<TRow> where TRow : IRowWriter
    {
        private readonly IWorkbookWriter<TSheet> _workbook;
        private readonly HashSet<string> _sheetNames = new(StringComparer.OrdinalIgnoreCase);

        /// <summary>
        /// Wraps an already-created, already-started workbook writer. Ownership of <paramref name="workbook"/>
        /// transfers to this instance, which disposes it when this instance is disposed.
        /// </summary>
        /// <param name="workbook">The started workbook writer to wrap.</param>
        public WorkbookRecordWriter(IWorkbookWriter<TSheet> workbook)
        {
            ArgumentNullException.ThrowIfNull(workbook);
            _workbook = workbook;
        }

        /// <summary>
        /// Writes a new sheet named <paramref name="sheetName"/> containing a header row followed by one
        /// row per item in <paramref name="records"/>.
        /// </summary>
        /// <typeparam name="T">The record type; its public readable, non-indexer properties (excluding
        /// any marked <c>[ExcelIgnore]</c>) each become a column.</typeparam>
        /// <param name="sheetName">The sheet's name; must be unique within this workbook.</param>
        /// <param name="records">The records to write, one row each, in enumeration order.</param>
        /// <param name="ct">A token to cancel the operation.</param>
        /// <exception cref="InvalidOperationException">A sheet named <paramref name="sheetName"/> already exists in this workbook.</exception>
        [RequiresUnreferencedCode("Record writing reflects over T's public properties, which trimming may remove.")]
        [RequiresDynamicCode("Record writing compiles the per-type column writer at runtime (Expression.Compile / MakeGenericMethod).")]
        public async ValueTask WriteSheetAsync<T>(string sheetName, IEnumerable<T> records, CancellationToken ct = default)
        {
            ArgumentNullException.ThrowIfNull(records);
            TSheet sheet = BeginSheet(sheetName);
            await using (sheet.ConfigureAwait(false))
            {
                await sheet.StartAsync(ct).ConfigureAwait(false);
                await WriteHeaderAsync<T>(sheet, ct).ConfigureAwait(false);
                await sheet.WriteRecordsAsync(records, RecordColumns<T>.WriteRow, ct).ConfigureAwait(false);
                await sheet.EndAsync(ct).ConfigureAwait(false);
            }
        }

        /// <summary>
        /// Writes a new sheet named <paramref name="sheetName"/> containing a header row followed by one
        /// row per item produced by <paramref name="records"/>.
        /// </summary>
        /// <typeparam name="T">The record type; its public readable, non-indexer properties (excluding
        /// any marked <c>[ExcelIgnore]</c>) each become a column.</typeparam>
        /// <param name="sheetName">The sheet's name; must be unique within this workbook.</param>
        /// <param name="records">The records to write, one row each, in enumeration order.</param>
        /// <param name="ct">A token to cancel the operation, and passed to the source enumerable.</param>
        /// <exception cref="InvalidOperationException">A sheet named <paramref name="sheetName"/> already exists in this workbook.</exception>
        [RequiresUnreferencedCode("Record writing reflects over T's public properties, which trimming may remove.")]
        [RequiresDynamicCode("Record writing compiles the per-type column writer at runtime (Expression.Compile / MakeGenericMethod).")]
        public async ValueTask WriteSheetAsync<T>(string sheetName, IAsyncEnumerable<T> records, CancellationToken ct = default)
        {
            ArgumentNullException.ThrowIfNull(records);
            TSheet sheet = BeginSheet(sheetName);
            await using (sheet.ConfigureAwait(false))
            {
                await sheet.StartAsync(ct).ConfigureAwait(false);
                await WriteHeaderAsync<T>(sheet, ct).ConfigureAwait(false);
                await sheet.WriteRecordsAsync(records, RecordColumns<T>.WriteRow, ct).ConfigureAwait(false);
                await sheet.EndAsync(ct).ConfigureAwait(false);
            }
        }

        private TSheet BeginSheet(string sheetName)
        {
            ArgumentNullException.ThrowIfNull(sheetName);
            if (!_sheetNames.Add(sheetName))
            {
                throw new InvalidOperationException($"A sheet named '{sheetName}' already exists in this workbook.");
            }
            return _workbook.AddSheet(sheetName);
        }

        [RequiresUnreferencedCode("Record writing reflects over T's public properties, which trimming may remove.")]
        [RequiresDynamicCode("Record writing compiles the per-type column writer at runtime (Expression.Compile / MakeGenericMethod).")]
        private static async ValueTask WriteHeaderAsync<T>(TSheet sheet, CancellationToken ct)
        {
            TRow row = await sheet.StartRowAsync(ct).ConfigureAwait(false);
            await using (row.ConfigureAwait(false))
            {
                foreach (string header in RecordColumns<T>.Headers)
                {
                    row.Write(header);
                }
            }
        }

        /// <summary>Finalizes and disposes the underlying workbook writer, completing the workbook.</summary>
        [SuppressMessage("IDisposableAnalyzers.Correctness", "IDISP007:Don't dispose injected",
            Justification = "The RecordWriter factories transfer ownership of the workbook to this wrapper, which finalizes it here.")]
        public ValueTask DisposeAsync()
        {
            // Each workbook writer's DisposeAsync ends the workbook (finalizing all sheets) when started.
            return _workbook.DisposeAsync();
        }
    }

    /// <summary>
    /// Format-specific factories that create the underlying low-level workbook writer, start it, and wrap
    /// it in a <see cref="WorkbookRecordWriter{TSheet,TRow}"/>, so callers never need to name the writer's
    /// type parameters themselves.
    /// </summary>
    public static class RecordWriter
    {
        /// <summary>Creates and starts a record writer that produces an XLSX workbook.</summary>
        /// <param name="stream">The destination stream; must support writing.</param>
        /// <param name="leaveOpen">If <see langword="true"/>, <paramref name="stream"/> is left open when the returned writer is disposed.</param>
        /// <param name="compression">The ZIP compression level used for the XLSX package's entries.</param>
        /// <param name="useSharedStrings">If <see langword="true"/>, text cells are deduplicated into the shared-strings table instead of written inline.</param>
        /// <param name="prefetchWrite">If <see langword="true"/>, each sheet's deflate runs on a background thread instead of the calling thread (see <see cref="XlsxWorkbookWriter.CreateAsync"/>).</param>
        /// <param name="ct">A token to cancel the operation.</param>
        /// <returns>A started record writer ready to accept sheets.</returns>
        public static async ValueTask<WorkbookRecordWriter<XlsxSheetWriter, XlsxRowWriter>> CreateXlsxAsync(
            Stream stream, bool leaveOpen = false, CompressionLevel compression = CompressionLevel.Fastest,
            bool useSharedStrings = false, bool prefetchWrite = false, CancellationToken ct = default)
        {
            var workbook = await XlsxWorkbookWriter.CreateAsync(stream, leaveOpen, compression, useSharedStrings, prefetchWrite, ct).ConfigureAwait(false);
            await workbook.StartAsync(ct).ConfigureAwait(false);
            return new WorkbookRecordWriter<XlsxSheetWriter, XlsxRowWriter>(workbook);
        }

        /// <summary>Creates and starts a record writer that produces an XLSB workbook.</summary>
        /// <param name="stream">The destination stream; must support writing.</param>
        /// <param name="leaveOpen">If <see langword="true"/>, <paramref name="stream"/> is left open when the returned writer is disposed.</param>
        /// <param name="date1904">If <see langword="true"/>, dates are serialized using the 1904 date system instead of the default 1900 system.</param>
        /// <param name="compression">The ZIP compression level used for the XLSB package's entries.</param>
        /// <param name="useSharedStrings">If <see langword="true"/>, text cells are deduplicated into the shared-strings table instead of written inline.</param>
        /// <param name="prefetchWrite">If <see langword="true"/>, each sheet's deflate runs on a background thread instead of the calling thread (see <see cref="XlsbWorkbookWriter.CreateAsync"/>).</param>
        /// <param name="ct">A token to cancel the operation.</param>
        /// <returns>A started record writer ready to accept sheets.</returns>
        public static async ValueTask<WorkbookRecordWriter<XlsbSheetWriter, XlsbRowWriter>> CreateXlsbAsync(
            Stream stream, bool leaveOpen = false, bool date1904 = false,
            CompressionLevel compression = CompressionLevel.Fastest, bool useSharedStrings = false,
            bool prefetchWrite = false, CancellationToken ct = default)
        {
            var workbook = await XlsbWorkbookWriter.CreateAsync(stream, leaveOpen, date1904, compression, useSharedStrings, prefetchWrite, ct).ConfigureAwait(false);
            await workbook.StartAsync(ct).ConfigureAwait(false);
            return new WorkbookRecordWriter<XlsbSheetWriter, XlsbRowWriter>(workbook);
        }

        /// <summary>
        /// Creates and starts a record writer that produces a CSV file. Supports only a single sheet,
        /// since a CSV file is inherently one sheet. The returned writer is still <see cref="IAsyncDisposable"/>.
        /// </summary>
        /// <param name="stream">The destination stream; must support writing.</param>
        /// <param name="leaveOpen">If <see langword="true"/>, <paramref name="stream"/> is left open when the returned writer is disposed.</param>
        /// <param name="options">The delimiter/quote character to use; defaults to <see cref="CsvWriterOptions.Default"/> if <see langword="null"/>.</param>
        /// <param name="ct">A token to cancel the operation.</param>
        /// <returns>A started record writer ready to accept its single sheet.</returns>
        [SuppressMessage("Reliability", "CA2000:Dispose objects before losing scope",
            Justification = "Ownership of the workbook transfers to WorkbookRecordWriter, which disposes it.")]
        public static async ValueTask<WorkbookRecordWriter<CsvSheetWriter, CsvRowWriter>> CreateCsvAsync(
            Stream stream, bool leaveOpen = false, CsvWriterOptions? options = null, CancellationToken ct = default)
        {
            var workbook = CsvWorkbookWriter.Create(stream, leaveOpen, options);
            await workbook.StartAsync(ct).ConfigureAwait(false);
            return new WorkbookRecordWriter<CsvSheetWriter, CsvRowWriter>(workbook);
        }

        /// <summary>Creates and starts a record writer that produces a legacy XLS (BIFF8) workbook.</summary>
        /// <param name="stream">The destination stream; must support writing.</param>
        /// <param name="leaveOpen">If <see langword="true"/>, <paramref name="stream"/> is left open when the returned writer is disposed.</param>
        /// <param name="date1904">If <see langword="true"/>, dates are serialized using the 1904 date system instead of the default 1900 system.</param>
        /// <param name="ct">A token to cancel the operation.</param>
        /// <returns>A started record writer ready to accept sheets.</returns>
        [SuppressMessage("Reliability", "CA2000:Dispose objects before losing scope",
            Justification = "Ownership of the workbook transfers to WorkbookRecordWriter, which disposes it.")]
        public static async ValueTask<WorkbookRecordWriter<XlsSheetWriter, XlsRowWriter>> CreateXlsAsync(
            Stream stream, bool leaveOpen = false, bool date1904 = false, CancellationToken ct = default)
        {
            var workbook = XlsWorkbookWriter.Create(stream, leaveOpen, date1904);
            await workbook.StartAsync(ct).ConfigureAwait(false);
            return new WorkbookRecordWriter<XlsSheetWriter, XlsRowWriter>(workbook);
        }
    }

    // Per-type column plan, cached per (T, TRow) via the nested Plan<TRow>. Compiled against the
    // concrete TRow rather than IRowWriter, so each call resolves directly to that sealed class's
    // non-virtual method instead of an interface dispatch. Non-numeric/non-primitive properties fall
    // back to ToString() text, since Write<U> only produces valid numeric cells.
    [RequiresUnreferencedCode("Record writing reflects over T's public properties, which trimming may remove.")]
    [RequiresDynamicCode("Record writing compiles the per-type column writer at runtime (Expression.Compile / MakeGenericMethod).")]
    [SuppressMessage("Major Code Smell", "S2743:Static fields should not be used in generic types",
        Justification = "The per-closed-type static IS the design: headers/property plan are cached once per T, not shared across different T.")]
    internal static class RecordColumns<T>
    {
        private static readonly PropertyInfo[] _props = FilterProperties();
        internal static string[] Headers { get; } = BuildHeaders(_props);

        internal static void WriteRow<TRow>(TRow row, T record) where TRow : IRowWriter
        {
            Plan<TRow>.Write(row, record);
        }

        private static PropertyInfo[] FilterProperties()
        {
            return [.. typeof(T).GetProperties(BindingFlags.Public | BindingFlags.Instance)
                .Where(prop => prop.GetGetMethod() is not null && prop.GetIndexParameters().Length == 0
                    && !Attribute.IsDefined(prop, typeof(ExcelIgnoreAttribute)))];
        }

        private static string[] BuildHeaders(PropertyInfo[] props)
        {
            var headers = new string[props.Length];
            for (int i = 0; i < props.Length; i++)
            {
                // First [ExcelColumn] wins (the attribute allows aliases); fall back to the property name.
                ExcelColumnAttribute? attr = props[i].GetCustomAttributes<ExcelColumnAttribute>().FirstOrDefault();
                headers[i] = attr?.Name ?? props[i].Name;
            }
            return headers;
        }

        // Keyed by TRow: one compiled Action<TRow, T> per concrete row-writer type actually used with T.
        [RequiresUnreferencedCode("Record writing reflects over T's public properties, which trimming may remove.")]
        [RequiresDynamicCode("Record writing compiles the per-type column writer at runtime (Expression.Compile / MakeGenericMethod).")]
        private static class Plan<TRow> where TRow : IRowWriter
        {
            internal static readonly Action<TRow, T> Write = Build();
            // Null-safe ToString() for a non-string, non-numeric property, so it lands in a text cell.
            private static Expression ToStringExpression(Expression value, Type pt)
            {
                MethodInfo toString = typeof(object).GetMethod(nameof(ToString), Type.EmptyTypes)!;
                if (pt.IsValueType && Nullable.GetUnderlyingType(pt) is null)
                {
                    // Never null; box once for the virtual ToString call (rare path).
                    return Expression.Call(Expression.Convert(value, typeof(object)), toString);
                }
                Expression boxed = Expression.Convert(value, typeof(object));
                return Expression.Condition(
                    Expression.NotEqual(boxed, Expression.Constant(null, typeof(object))),
                    Expression.Call(boxed, toString),
                    Expression.Constant(null, typeof(string)));
            }
            // `instances` caches by converter type so one used on multiple properties is built once.
            private static MethodCallExpression? TryBuildConverterWrite(
                PropertyInfo prop, ParameterExpression rowParam, Expression value, Dictionary<Type, object> instances)
            {
                ExcelConverterAttribute? converter = prop.GetCustomAttribute<ExcelConverterAttribute>();
                if (converter is null)
                {
                    return null;
                }
                Type writerInterface = typeof(IExcelCellWriter<>).MakeGenericType(prop.PropertyType);
                if (!writerInterface.IsAssignableFrom(converter.ConverterType))
                {
                    return null;
                }
                if (!instances.TryGetValue(converter.ConverterType, out object? instance))
                {
                    instance = Activator.CreateInstance(converter.ConverterType)
                        ?? throw new InvalidOperationException($"Converter '{converter.ConverterType}' could not be instantiated.");
                    instances.Add(converter.ConverterType, instance);
                }
                MethodInfo writeMethod = writerInterface.GetMethod(nameof(IExcelCellWriter<>.Write))!;
                return Expression.Call(Expression.Constant(instance, writerInterface), writeMethod, rowParam, value);
            }
            private static Action<TRow, T> Build()
            {
                ParameterExpression rowParam = Expression.Parameter(typeof(TRow), "row");
                ParameterExpression recParam = Expression.Parameter(typeof(T), "record");
                var body = new List<Expression>(_props.Length);
                var converterInstances = new Dictionary<Type, object>();

                foreach (PropertyInfo prop in _props)
                {
                    Expression value = Expression.Property(recParam, prop);
                    Expression? converterCall = TryBuildConverterWrite(prop, rowParam, value, converterInstances);
                    if (converterCall is not null)
                    {
                        body.Add(converterCall);
                        continue;
                    }
                    MethodInfo write = RowWriteMethods<TRow>.Select(prop.PropertyType, out bool asString);
                    if (asString)
                    {
                        value = ToStringExpression(value, prop.PropertyType);
                    }
                    body.Add(Expression.Call(rowParam, write, value));
                }

                return body.Count == 0 ? static (_, _) => { } : Expression.Lambda<Action<TRow, T>>(Expression.Block(body), rowParam, recParam).Compile();
            }
        }
    }

    // Reflection resolved once per concrete TRow: the Write overloads it declares, plus the set of
    // numeric property types that map to Write<U>.
    [RequiresUnreferencedCode("Record writing reflects over TRow's Write overloads, which trimming may remove.")]
    [RequiresDynamicCode("Record writing dispatches through MakeGenericMethod for numeric column types.")]
    [SuppressMessage("Major Code Smell", "S2743:Static fields should not be used in generic types",
        Justification = "The per-closed-type static IS the design: the resolved MethodInfo set is cached once per concrete TRow, not shared across different TRow.")]
    internal static class RowWriteMethods<TRow> where TRow : IRowWriter
    {
        private readonly record struct MethodInfoSet(MethodInfo Str, MethodInfo Bool, MethodInfo BoolN, MethodInfo Date,
                                                     MethodInfo DateN, MethodInfo DateOnly, MethodInfo DateOnlyN,
                                                     MethodInfo TimeOnly, MethodInfo TimeOnlyN,
                                                     MethodInfo Generic, MethodInfo GenericN);
        private static readonly HashSet<Type> Numeric =
        [
            typeof(byte), typeof(sbyte), typeof(short), typeof(ushort), typeof(int), typeof(uint),
            typeof(long), typeof(ulong), typeof(float), typeof(double), typeof(decimal),
        ];

        private static readonly MethodInfoSet M = Resolve();

        private static MethodInfoSet Resolve()
        {
            MethodInfo? str = null, boolean = null, booleanN = null, date = null, dateN = null,
                dateOnly = null, dateOnlyN = null, timeOnly = null, timeOnlyN = null, generic = null, genericN = null;
            foreach (MethodInfo m in typeof(TRow).GetMethods())
            {
                if (!string.Equals(m.Name, nameof(IRowWriter.Write), StringComparison.Ordinal))
                {
                    continue;
                }
                Type p = m.GetParameters()[0].ParameterType;
                if (m.IsGenericMethodDefinition)
                {
                    if (p == m.GetGenericArguments()[0]) { generic = m; }
                    else { genericN = m; } // Write<U>(U?)
                }
                else if (p == typeof(string)) { str = m; }
                else if (p == typeof(bool)) { boolean = m; }
                else if (p == typeof(bool?)) { booleanN = m; }
                else if (p == typeof(DateTime)) { date = m; }
                else if (p == typeof(DateTime?)) { dateN = m; }
                else if (p == typeof(DateOnly)) { dateOnly = m; }
                else if (p == typeof(DateOnly?)) { dateOnlyN = m; }
                else if (p == typeof(TimeOnly)) { timeOnly = m; }
                else if (p == typeof(TimeOnly?)) { timeOnlyN = m; }
            }
            return new(str!, boolean!, booleanN!, date!, dateN!, dateOnly!, dateOnlyN!, timeOnly!, timeOnlyN!, generic!, genericN!);
        }

        // Picks the concrete TRow.Write overload for a property type; asString means the caller must
        // first convert the value to a string (the non-numeric fallback).
        internal static MethodInfo Select(Type pt, out bool asString)
        {
            asString = false;
            if (pt == typeof(string)) { return M.Str; }
            if (pt == typeof(bool)) { return M.Bool; }
            if (pt == typeof(bool?)) { return M.BoolN; }
            if (pt == typeof(DateTime)) { return M.Date; }
            if (pt == typeof(DateTime?)) { return M.DateN; }
            if (pt == typeof(DateOnly)) { return M.DateOnly; }
            if (pt == typeof(DateOnly?)) { return M.DateOnlyN; }
            if (pt == typeof(TimeOnly)) { return M.TimeOnly; }
            if (pt == typeof(TimeOnly?)) { return M.TimeOnlyN; }
            Type? underlying = Nullable.GetUnderlyingType(pt);
            if (underlying is null && Numeric.Contains(pt)) { return M.Generic.MakeGenericMethod(pt); }
            if (underlying is not null && Numeric.Contains(underlying)) { return M.GenericN.MakeGenericMethod(underlying); }
            asString = true;
            return M.Str;
        }
    }
}
