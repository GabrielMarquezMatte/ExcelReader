using System.Diagnostics.CodeAnalysis;
using System.IO.Compression;
using System.Linq.Expressions;
using System.Reflection;
using ExcelReader.Core.Parser;

namespace ExcelReader.Core.Writer
{
    // High-level record writer: writes a header row (from [ExcelColumn] or the property name) followed
    // by one row per record, mapping each public readable property to a column. Each call targets a new
    // sheet, so one workbook can hold sheets of different record types. Generic over the low-level
    // IWorkbookWriter/ISheetWriter/IRowWriter interfaces, so it works for XLSX, XLSB and XLS alike.
    // Round-trips through ExcelParser<T> because headers match property names/aliases (column order is
    // irrelevant there). Use the RecordWriter.Create* factories rather than constructing directly.
    public sealed class WorkbookRecordWriter<TSheet, TRow> : IAsyncDisposable where TSheet : ISheetWriter<TRow> where TRow : IRowWriter
    {
        private readonly IWorkbookWriter<TSheet> _workbook;
        private readonly HashSet<string> _sheetNames = new(StringComparer.OrdinalIgnoreCase);

        public WorkbookRecordWriter(IWorkbookWriter<TSheet> workbook)
        {
            ArgumentNullException.ThrowIfNull(workbook);
            _workbook = workbook;
        }

        [RequiresUnreferencedCode("Record writing reflects over T's public properties, which trimming may remove.")]
        [RequiresDynamicCode("Record writing compiles the per-type column writer at runtime (Expression.Compile / MakeGenericMethod).")]
        public async ValueTask WriteSheetAsync<T>(string sheetName, IEnumerable<T> records, CancellationToken ct = default)
        {
            ArgumentNullException.ThrowIfNull(records);
            TSheet sheet = BeginSheet(sheetName);
            await sheet.StartAsync(ct).ConfigureAwait(false);
            await WriteHeaderAsync<T>(sheet, ct).ConfigureAwait(false);
            await sheet.WriteRecordsAsync(records, RecordColumns<T>.WriteRow, ct).ConfigureAwait(false);
            await sheet.EndAsync(ct).ConfigureAwait(false);
        }

        [RequiresUnreferencedCode("Record writing reflects over T's public properties, which trimming may remove.")]
        [RequiresDynamicCode("Record writing compiles the per-type column writer at runtime (Expression.Compile / MakeGenericMethod).")]
        public async ValueTask WriteSheetAsync<T>(string sheetName, IAsyncEnumerable<T> records, CancellationToken ct = default)
        {
            ArgumentNullException.ThrowIfNull(records);
            TSheet sheet = BeginSheet(sheetName);
            await sheet.StartAsync(ct).ConfigureAwait(false);
            await WriteHeaderAsync<T>(sheet, ct).ConfigureAwait(false);
            await sheet.WriteRecordsAsync(records, RecordColumns<T>.WriteRow, ct).ConfigureAwait(false);
            await sheet.EndAsync(ct).ConfigureAwait(false);
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

        [SuppressMessage("IDisposableAnalyzers.Correctness", "IDISP007:Don't dispose injected",
            Justification = "The RecordWriter factories transfer ownership of the workbook to this wrapper, which finalizes it here.")]
        public ValueTask DisposeAsync()
        {
            // Each workbook writer's DisposeAsync ends the workbook (finalizing all sheets) when started.
            return _workbook.DisposeAsync();
        }
    }

    // Format-specific factories for the generic record writer: they create the low-level workbook,
    // start it, and hand it to WorkbookRecordWriter so callers never touch the type parameters.
    public static class RecordWriter
    {
        public static async ValueTask<WorkbookRecordWriter<SheetWriter, RowWriter>> CreateXlsxAsync(
            Stream stream, bool leaveOpen = false, CompressionLevel compression = CompressionLevel.Fastest,
            bool useSharedStrings = false, CancellationToken ct = default)
        {
            var workbook = await WorkbookWriter.CreateAsync(stream, leaveOpen, compression, ct, useSharedStrings).ConfigureAwait(false);
            await workbook.StartAsync(ct).ConfigureAwait(false);
            return new WorkbookRecordWriter<SheetWriter, RowWriter>(workbook);
        }

        public static async ValueTask<WorkbookRecordWriter<XlsbSheetWriter, XlsbRowWriter>> CreateXlsbAsync(
            Stream stream, bool leaveOpen = false, bool date1904 = false,
            CompressionLevel compression = CompressionLevel.Fastest, bool useSharedStrings = false,
            CancellationToken ct = default)
        {
            var workbook = await XlsbWorkbookWriter.CreateAsync(stream, leaveOpen, date1904, compression, ct, useSharedStrings).ConfigureAwait(false);
            await workbook.StartAsync(ct).ConfigureAwait(false);
            return new WorkbookRecordWriter<XlsbSheetWriter, XlsbRowWriter>(workbook);
        }

        // CSV has no async setup (no ZIP/BIFF headers to write), so this is synchronous — the returned
        // writer is still IAsyncDisposable. Only a single sheet is supported (a CSV file is one sheet).
        [SuppressMessage("Reliability", "CA2000:Dispose objects before losing scope",
            Justification = "Ownership of the workbook transfers to WorkbookRecordWriter, which disposes it.")]
        public static WorkbookRecordWriter<CsvSheetWriter, CsvRowWriter> CreateCsv(
            Stream stream, bool leaveOpen = false, CsvWriterOptions? options = null)
        {
            var workbook = CsvWorkbookWriter.Create(stream, leaveOpen, options);
            return new WorkbookRecordWriter<CsvSheetWriter, CsvRowWriter>(workbook);
        }

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

    // Per-type column plan. Headers depend only on T; the write delegate additionally depends on the
    // concrete TRow (RowWriter/XlsbRowWriter/XlsRowWriter), cached per (T, TRow) via the nested Plan<TRow>.
    // Compiling against the concrete TRow (instead of the IRowWriter interface) lets each Expression.Call
    // resolve directly to that sealed class's non-virtual method, so the JIT emits a direct call per cell
    // instead of an interface dispatch. Numeric properties go to the generic Write<U> (a number cell)
    // everything non-numeric/non-primitive is written as its ToString() text, since Write<U> only
    // produces valid numeric cells.
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
        private static class Plan<TRow> where TRow : IRowWriter
        {
            internal static readonly Action<TRow, T> Write = Build();
            // Produces a string? expression for a non-string, non-numeric property (Guid, enum, char, DateOnly,
            // custom types): null-safe ToString() so it lands in a text cell rather than a corrupt number cell.
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
            // If the property carries [ExcelConverter(T)] and T implements IExcelCellWriter<propType>, returns
            // a call to that converter's Write; otherwise null so the caller falls back to type-based
            // routing (a read-only converter simply gets no write side). `instances` caches by converter
            // type so a converter used on multiple properties of the same (T, TRow) is built only once.
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
                    // A [ExcelConverter] whose type also implements IExcelCellWriter<propType> owns the write
                    // (round-trips a custom type written here and read back via its IExcelCellConverter side).
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

    // Reflection resolved once per concrete TRow (RowWriter/XlsbRowWriter/XlsRowWriter — at most a
    // handful of instantiations for the whole process): the Write overloads declared on that concrete
    // type, plus the set of numeric property types that map to Write<U>.
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
