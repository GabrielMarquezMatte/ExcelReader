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

        public async ValueTask WriteSheetAsync<T>(string sheetName, IEnumerable<T> records, CancellationToken ct = default)
        {
            ArgumentNullException.ThrowIfNull(records);
            TSheet sheet = BeginSheet(sheetName);
            await sheet.StartAsync(ct).ConfigureAwait(false);
            await WriteHeaderAsync<T>(sheet, ct).ConfigureAwait(false);
            await sheet.WriteRecordsAsync(records, RecordColumns<T>.WriteRow, ct).ConfigureAwait(false);
            await sheet.EndAsync(ct).ConfigureAwait(false);
        }

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

    // Per-type column plan, built once and cached. The writer is an expression-tree-compiled
    // Action<IRowWriter, T>: each property is read strongly-typed and routed to the matching IRowWriter
    // overload with no boxing and no per-cell type switch. Numeric properties go to the generic Write<U>
    // (a number cell); everything non-numeric/non-primitive is written as its ToString() text, since
    // Write<U> only produces valid numeric cells.
    internal static class RecordColumns<T>
    {
        // Single field whose type mentions T, so the per-close-type caching is deliberate (avoids S2743).
        private static readonly (string[] Headers, Action<IRowWriter, T> Write) _plan = Build();
        internal static string[] Headers => _plan.Headers;
        // Write targets IRowWriter (every row writer is a reference type), so a TRow argument converts with no box.
        internal static void WriteRow<TRow>(TRow row, T record) where TRow : IRowWriter
        {
            _plan.Write(row, record);
        }

        private static (string[], Action<IRowWriter, T>) Build()
        {
            PropertyInfo[] props = typeof(T).GetProperties(BindingFlags.Public | BindingFlags.Instance);
            var headers = new List<string>(props.Length);
            ParameterExpression rowParam = Expression.Parameter(typeof(IRowWriter), "row");
            ParameterExpression recParam = Expression.Parameter(typeof(T), "record");
            var body = new List<Expression>(props.Length);

            foreach (PropertyInfo prop in props)
            {
                if (prop.GetGetMethod() is null || prop.GetIndexParameters().Length > 0)
                {
                    continue;
                }
                // First [ExcelColumn] wins (the attribute allows aliases); fall back to the property name.
                ExcelColumnAttribute? attr = prop.GetCustomAttributes<ExcelColumnAttribute>().FirstOrDefault();
                headers.Add(attr?.Name ?? prop.Name);

                Expression value = Expression.Property(recParam, prop);
                MethodInfo write = RowWriteMethods.Select(prop.PropertyType, out bool asString);
                if (asString)
                {
                    value = ToStringExpression(value, prop.PropertyType);
                }
                body.Add(Expression.Call(rowParam, write, value));
            }

            var writeRow = body.Count == 0 ? static (_, _) => { } : Expression.Lambda<Action<IRowWriter, T>>(Expression.Block(body), rowParam, recParam).Compile();
            return ([.. headers], writeRow);
        }

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
    }

    // T-independent reflection resolved once for the whole process (not per closed RecordColumns<T>):
    // the IRowWriter.Write overloads and the set of numeric property types that map to Write<U>.
    internal static class RowWriteMethods
    {
        private readonly record struct MethodInfoSet(MethodInfo Str, MethodInfo Bool, MethodInfo BoolN, MethodInfo Date,
                                                     MethodInfo DateN, MethodInfo DateOnly, MethodInfo DateOnlyN,
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
                dateOnly = null, dateOnlyN = null, generic = null, genericN = null;
            foreach (MethodInfo m in typeof(IRowWriter).GetMethods())
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
            }
            return new(str!, boolean!, booleanN!, date!, dateN!, dateOnly!, dateOnlyN!, generic!, genericN!);
        }

        // Picks the IRowWriter.Write overload for a property type; asString means the caller must first
        // convert the value to a string (the non-numeric fallback).
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
            Type? underlying = Nullable.GetUnderlyingType(pt);
            if (underlying is null && Numeric.Contains(pt)) { return M.Generic.MakeGenericMethod(pt); }
            if (underlying is not null && Numeric.Contains(underlying)) { return M.GenericN.MakeGenericMethod(underlying); }
            asString = true;
            return M.Str;
        }
    }
}
