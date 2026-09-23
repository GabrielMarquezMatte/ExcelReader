using System.Diagnostics.CodeAnalysis;
using System.Globalization;
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
    public sealed class WorkbookRecordWriter<TSheet, TRow> : WorkbookRecordWriterBase<TSheet, TRow>
        where TSheet : ISheetWriter<TRow>
        where TRow : IRowWriter
    {
        /// <summary>
        /// Wraps an already-created workbook writer. Ownership of <paramref name="workbook"/>
        /// transfers to this instance, which disposes it when this instance is disposed.
        /// </summary>
        /// <param name="workbook">The workbook writer to wrap.</param>
        public WorkbookRecordWriter(IWorkbookWriter<TSheet> workbook)
            : base(workbook)
        {
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
        public ValueTask WriteSheetAsync<T>(string sheetName, IEnumerable<T> records, CancellationToken ct = default)
        {
            return WriteSheetCoreAsync(sheetName, records, RecordColumns<T>.Headers, RecordColumns<T>.WriteRow, ct);
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
        public ValueTask WriteSheetAsync<T>(string sheetName, IAsyncEnumerable<T> records, CancellationToken ct = default)
        {
            return WriteSheetCoreAsync(sheetName, records, RecordColumns<T>.Headers, RecordColumns<T>.WriteRow, ct);
        }
    }

    /// <summary>
    /// Format-specific factories that create the underlying low-level workbook writer and wrap
    /// it in a <see cref="WorkbookRecordWriter{TSheet,TRow}"/>, so callers never need to name the writer's
    /// type parameters themselves.
    /// </summary>
    public static class RecordWriter
    {
        /// <summary>Creates a record writer that produces an XLSX workbook.</summary>
        /// <param name="stream">The destination stream; must support writing.</param>
        /// <param name="leaveOpen">If <see langword="true"/>, <paramref name="stream"/> is left open when the returned writer is disposed.</param>
        /// <param name="options">Compression, shared-string and background-deflate settings. Defaults to <see cref="XlsxWriterOptions.Default"/>.</param>
        /// <returns>A record writer ready to accept sheets.</returns>
        public static WorkbookRecordWriter<XlsxSheetWriter, XlsxRowWriter> CreateXlsx(
            Stream stream, bool leaveOpen = false, XlsxWriterOptions? options = null)
        {
            return new WorkbookRecordWriter<XlsxSheetWriter, XlsxRowWriter>(XlsxWorkbookWriter.Create(stream, leaveOpen, options));
        }

        /// <summary>Creates a record writer that produces an XLSB workbook.</summary>
        /// <param name="stream">The destination stream; must support writing.</param>
        /// <param name="leaveOpen">If <see langword="true"/>, <paramref name="stream"/> is left open when the returned writer is disposed.</param>
        /// <param name="options">Date system, compression, shared-string and background-deflate settings. Defaults to <see cref="XlsbWriterOptions.Default"/>.</param>
        /// <returns>A record writer ready to accept sheets.</returns>
        public static WorkbookRecordWriter<XlsbSheetWriter, XlsbRowWriter> CreateXlsb(
            Stream stream, bool leaveOpen = false, XlsbWriterOptions? options = null)
        {
            return new WorkbookRecordWriter<XlsbSheetWriter, XlsbRowWriter>(XlsbWorkbookWriter.Create(stream, leaveOpen, options));
        }

        /// <summary>
        /// Creates a record writer that produces a CSV file. Supports only a single sheet, since a CSV
        /// file is inherently one sheet.
        /// </summary>
        /// <param name="stream">The destination stream; must support writing.</param>
        /// <param name="leaveOpen">If <see langword="true"/>, <paramref name="stream"/> is left open when the returned writer is disposed.</param>
        /// <param name="options">The delimiter/quote character to use; defaults to <see cref="CsvWriterOptions.Default"/> if <see langword="null"/>.</param>
        /// <returns>A record writer ready to accept its single sheet.</returns>
        public static WorkbookRecordWriter<CsvSheetWriter, CsvRowWriter> CreateCsv(
            Stream stream, bool leaveOpen = false, CsvWriterOptions? options = null)
        {
            return new WorkbookRecordWriter<CsvSheetWriter, CsvRowWriter>(CsvWorkbookWriter.Create(stream, leaveOpen, options));
        }

        /// <summary>Creates a record writer that produces a legacy XLS (BIFF8) workbook.</summary>
        /// <param name="stream">The destination stream; must support writing.</param>
        /// <param name="leaveOpen">If <see langword="true"/>, <paramref name="stream"/> is left open when the returned writer is disposed.</param>
        /// <param name="date1904">If <see langword="true"/>, dates are serialized using the 1904 date system instead of the default 1900 system.</param>
        /// <returns>A record writer ready to accept sheets.</returns>
        public static WorkbookRecordWriter<XlsSheetWriter, XlsRowWriter> CreateXls(
            Stream stream, bool leaveOpen = false, bool date1904 = false)
        {
            return new WorkbookRecordWriter<XlsSheetWriter, XlsRowWriter>(XlsWorkbookWriter.Create(stream, leaveOpen, date1904));
        }
    }

    [RequiresUnreferencedCode("Record writing reflects over T's public properties, which trimming may remove.")]
    [RequiresDynamicCode("Record writing compiles the per-type column writer at runtime (Expression.Compile / MakeGenericMethod).")]
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
                ExcelColumnAttribute? attr = props[i].GetCustomAttributes<ExcelColumnAttribute>().FirstOrDefault();
                headers[i] = attr?.Name ?? props[i].Name;
            }
            return headers;
        }

        [RequiresUnreferencedCode("Record writing reflects over T's public properties, which trimming may remove.")]
        [RequiresDynamicCode("Record writing compiles the per-type column writer at runtime (Expression.Compile / MakeGenericMethod).")]
        private static class Plan<TRow> where TRow : IRowWriter
        {
            internal static readonly Action<TRow, T> Write = Build();
            private static Expression ToStringExpression(Expression value, Type pt)
            {
                Type? underlying = Nullable.GetUnderlyingType(pt);
                Type core = underlying ?? pt;
                if (typeof(IFormattable).IsAssignableFrom(core))
                {
                    string helper = underlying is not null ? nameof(InvariantText.FormatNullable)
                        : core.IsValueType ? nameof(InvariantText.FormatValue)
                        : nameof(InvariantText.FormatReference);
                    return Expression.Call(typeof(InvariantText).GetMethod(helper, BindingFlags.NonPublic | BindingFlags.Static)!.MakeGenericMethod(core), value);
                }
                MethodInfo toString = typeof(object).GetMethod(nameof(ToString), Type.EmptyTypes)!;
                if (pt.IsValueType && Nullable.GetUnderlyingType(pt) is null)
                {
                    return Expression.Call(Expression.Convert(value, typeof(object)), toString);
                }
                Expression boxed = Expression.Convert(value, typeof(object));
                return Expression.Condition(
                    Expression.NotEqual(boxed, Expression.Constant(null, typeof(object))),
                    Expression.Call(boxed, toString),
                    Expression.Constant(null, typeof(string)));
            }
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

    [RequiresUnreferencedCode("Record writing reflects over TRow's Write overloads, which trimming may remove.")]
    [RequiresDynamicCode("Record writing dispatches through MakeGenericMethod for numeric column types.")]
    internal static class RowWriteMethods<TRow> where TRow : IRowWriter
    {
        private readonly record struct MethodInfoSet(MethodInfo Str, MethodInfo Bool, MethodInfo BoolN, MethodInfo Date,
                                                     MethodInfo DateN, MethodInfo DateOnly, MethodInfo DateOnlyN,
                                                     MethodInfo TimeOnly, MethodInfo TimeOnlyN,
                                                     MethodInfo Generic, MethodInfo GenericN);
        private static readonly HashSet<Type> Numeric =
        [
            typeof(byte), typeof(sbyte), typeof(short), typeof(ushort), typeof(int), typeof(uint),
            typeof(long), typeof(ulong), typeof(float), typeof(double), typeof(decimal), typeof(Half),
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
                    else { genericN = m; }
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

    /// <summary>
    /// Formats text-fallback columns with the invariant culture, which is what <see cref="ExcelParserConfig.Culture"/>
    /// defaults to, so a column written on any machine parses back.
    /// </summary>
    internal static class InvariantText
    {
        internal static string FormatValue<TValue>(TValue value)
            where TValue : struct, IFormattable
        {
            return value.ToString(typeof(TValue) == typeof(DateTimeOffset) ? "O" : null, CultureInfo.InvariantCulture);
        }

        internal static string? FormatNullable<TValue>(TValue? value)
            where TValue : struct, IFormattable
        {
            return value.HasValue ? FormatValue(value.GetValueOrDefault()) : null;
        }

        internal static string? FormatReference<TValue>(TValue? value)
            where TValue : class, IFormattable
        {
            return value?.ToString(null, CultureInfo.InvariantCulture);
        }
    }
}
