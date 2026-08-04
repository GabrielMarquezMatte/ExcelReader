using System.Runtime.InteropServices;

namespace ExcelReader.Core.Writer
{
    /// <summary>
    /// Accumulates a property-to-column write plan for <typeparamref name="T"/>, fed either by the
    /// source generator (feature A) or by a hand-written <see cref="IExcelRecordMap{T}"/> implementation.
    /// The write-side counterpart to <see cref="Parser.ExcelRowMapBuilder{T}"/>.
    /// </summary>
    /// <typeparam name="T">The record type being mapped.</typeparam>
    public sealed class ExcelRecordMapBuilder<T>
    {
        private readonly List<string> _headers = [];
        private readonly List<Action<IRowWriter, T>> _writers = [];

        /// <summary>Adds a column: its header text, and how to write one record's value for it.</summary>
        /// <param name="header">The column's header text.</param>
        /// <param name="write">Writes exactly one cell for <paramref name="header"/>'s column, from a record.</param>
        /// <returns>This builder, for chaining.</returns>
        public ExcelRecordMapBuilder<T> Column(string header, Action<IRowWriter, T> write)
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

        internal void WriteRow(IRowWriter row, T record)
        {
            foreach (ref readonly Action<IRowWriter, T> write in CollectionsMarshal.AsSpan(_writers))
            {
                write(row, record);
            }
        }
    }
}
