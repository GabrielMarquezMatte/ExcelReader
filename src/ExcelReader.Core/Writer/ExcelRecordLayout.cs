using System.Diagnostics.CodeAnalysis;
using ExcelReader.Core.Writer.Internal;

namespace ExcelReader.Core.Writer
{
    /// <summary>
    /// How <see cref="SheetWriterExtensions"/>' <c>WriteRecordsAsync</c> turns a <typeparamref name="T"/> into a
    /// header row and one row per record. Create one with <see cref="ExcelRecordLayout"/>:
    /// <see cref="ExcelRecordLayout.FromAttributes{T}"/> reflects over the record type, and
    /// <see cref="ExcelRecordLayout.Generated{T}"/> uses the <c>[ExcelSerializable]</c> source-generated map.
    /// </summary>
    /// <typeparam name="T">The record type.</typeparam>
    public abstract class ExcelRecordLayout<T>
    {
        private protected ExcelRecordLayout()
        {
        }

        internal abstract (string[] Headers, Action<TRow, T> WriteRow) Columns<TRow>()
            where TRow : IRowWriter;
    }

    /// <summary>Creates <see cref="ExcelRecordLayout{T}"/> instances.</summary>
    public static class ExcelRecordLayout
    {
        /// <summary>
        /// Creates a layout with one column per public readable, non-indexer property of <typeparamref name="T"/>
        /// (excluding any marked <c>[ExcelIgnore]</c>), headed by its <c>[ExcelColumn]</c> name or, failing
        /// that, the property name, found by reflection.
        /// </summary>
        /// <remarks>
        /// Not compatible with Native AOT, and trimming can remove the properties it reads. Use
        /// <see cref="Generated{T}"/> instead where that matters.
        /// </remarks>
        /// <typeparam name="T">The record type.</typeparam>
        /// <returns>The layout.</returns>
        [RequiresUnreferencedCode("Record writing reflects over T's public properties, which trimming may remove.")]
        [RequiresDynamicCode("Record writing compiles the per-type column writer at runtime (Expression.Compile / MakeGenericMethod).")]
        public static ExcelRecordLayout<T> FromAttributes<T>()
        {
            return Reflected<T>.Instance;
        }

        /// <summary>Creates a layout from the map the <c>[ExcelSerializable]</c> source generator emitted for <typeparamref name="T"/>. Trimming and AOT safe.</summary>
        /// <typeparam name="T">The record type; must implement <see cref="IExcelRecordMap{T}"/>.</typeparam>
        /// <returns>The layout.</returns>
        public static ExcelRecordLayout<T> Generated<T>()
            where T : IExcelRecordMap<T>
        {
            return Mapped<T>.Instance;
        }

        private sealed class Reflected<T> : ExcelRecordLayout<T>
        {
            internal static readonly Reflected<T> Instance = new();

            [UnconditionalSuppressMessage("Trimming", "IL2026", Justification = "Only reachable through FromAttributes, which carries RequiresUnreferencedCode.")]
            [UnconditionalSuppressMessage("AOT", "IL3050", Justification = "Only reachable through FromAttributes, which carries RequiresDynamicCode.")]
            internal override (string[] Headers, Action<TRow, T> WriteRow) Columns<TRow>()
            {
                return (RecordColumns<T>.Headers, RecordColumns<T>.WriteRow);
            }
        }

        private sealed class Mapped<T> : ExcelRecordLayout<T>
            where T : IExcelRecordMap<T>
        {
            internal static readonly Mapped<T> Instance = new();

            internal override (string[] Headers, Action<TRow, T> WriteRow) Columns<TRow>()
            {
                return (MappedRecordColumns<T, TRow>.Headers, MappedRecordColumns<T, TRow>.WriteRow);
            }
        }
    }

    internal static class MappedRecordColumns<T, TRow>
        where T : IExcelRecordMap<T>
        where TRow : IRowWriter
    {
        private static readonly ExcelRecordMapBuilder<T, TRow> _builder = Build();

        internal static string[] Headers { get; } = _builder.Headers();

        internal static void WriteRow(TRow row, T record)
        {
            _builder.WriteRow(row, record);
        }

        private static ExcelRecordMapBuilder<T, TRow> Build()
        {
            var builder = new ExcelRecordMapBuilder<T, TRow>();
            T.ConfigureExcelRecordMap(builder);
            return builder;
        }
    }
}
