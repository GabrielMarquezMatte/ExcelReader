using System.Runtime.InteropServices;

namespace ExcelReader.Core.Writer
{
    /// <summary>
    /// Accumulates a property-to-column write plan for <typeparamref name="T"/>, fed either by the
    /// source generator (feature A) or by a hand-written <see cref="IExcelRecordMap{T}"/> implementation.
    /// The write-side counterpart to <see cref="Parser.ExcelRowMapBuilder{T}"/>.
    /// </summary>
    /// <typeparam name="T">The record type being mapped.</typeparam>
    /// <typeparam name="TRow">
    /// The concrete row writer the columns write through. A type parameter rather than
    /// <see cref="IRowWriter"/> so each column action compiles against that sealed class and its cell
    /// writes resolve to it directly instead of dispatching through the interface once per cell — the
    /// same trade <c>RecordColumns{T}.Plan{TRow}</c> makes on the reflection side. The cost is one
    /// plan per (<typeparamref name="T"/>, <typeparamref name="TRow"/>) pair instead of one per
    /// <typeparamref name="T"/>: a record written to three formats configures its map three times.
    /// </typeparam>
    public sealed class ExcelRecordMapBuilder<T, TRow>
        where TRow : IRowWriter
    {
        private readonly List<string> _headers = [];
        private readonly List<Action<TRow, T>> _writers = [];

        /// <summary>Adds a column: its header text, and how to write one record's value for it.</summary>
        /// <param name="header">The column's header text.</param>
        /// <param name="write">Writes exactly one cell for <paramref name="header"/>'s column, from a record.</param>
        /// <returns>This builder, for chaining.</returns>
        public ExcelRecordMapBuilder<T, TRow> Column(string header, Action<TRow, T> write)
        {
            ArgumentNullException.ThrowIfNull(header);
            ArgumentNullException.ThrowIfNull(write);
            _headers.Add(header);
            _writers.Add(write);
            return this;
        }

        internal string[] Headers()
        {
            return [.. _headers];
        }

        internal void WriteRow(TRow row, T record)
        {
            foreach (ref readonly Action<TRow, T> write in CollectionsMarshal.AsSpan(_writers))
            {
                write(row, record);
            }
        }
    }
}
