using System.Collections;
using System.Diagnostics.CodeAnalysis;
using ExcelReader.Core.Reader;

namespace ExcelReader.Core.Parser.Internal
{
    /// <summary>Base class supplying the shared row-advancement loop for a synchronous, format-specific row enumerator.</summary>
    /// <typeparam name="T">The row model type each derived enumerator yields.</typeparam>
    /// <typeparam name="TRows">The concrete row enumerator type this instance drives.</typeparam>
    /// <remarks>
    /// Shared by every row-projecting <see cref="IEnumerable{T}"/> (Excel formats + CSV): loop until
    /// <c>Rows.MoveNext()</c> is exhausted, project each row via the format-specific <c>Project()</c>
    /// override, stopping early on <c>ProjectionStep.Stop</c>. <c>Project()</c> is the only thing that
    /// differs per format (Excel walks <c>Row.Cells</c> generically via <c>RowProjector&lt;T&gt;</c>; CSV
    /// binds fields by dense index via <c>CsvRowProjector&lt;T&gt;</c>) — see
    /// <c>ExcelEnumerable&lt;T,TReader,TEnumerator&gt;.Enumerator</c> and <c>CsvEnumerable&lt;T&gt;.Enumerator</c>.
    /// Public because it is the base class of those public nested types (a base class can never be less
    /// accessible than its derived type). Mirrored by <see cref="AsyncRowEnumerator{T, TReader, TRows}"/>
    /// for the async side.
    /// </remarks>
    [SuppressMessage("Design", "CA1063:Implement IDisposable correctly",
        Justification = "No unmanaged resources and no finalizer; every derived Enumerator is sealed and adds no disposal logic, so the full Dispose(bool) pattern buys nothing here.")]
    public abstract class SyncRowEnumerator<T, TRows> : IEnumerator<T>
        where T : allows ref struct
        where TRows : class, IExcelRowEnumerator
    {
        /// <summary>The underlying row cursor this enumerator advances.</summary>
        [SuppressMessage("Design", "CA1051:Do not declare visible instance fields",
            Justification = "Hot-path base class (MoveNext runs per row); a field avoids a property-call indirection in the tightest loop of the library.")]
        protected readonly TRows Rows;

        /// <summary>Initializes the base enumerator with the row cursor it will drive.</summary>
        /// <param name="rows">The row enumerator to advance and project from.</param>
        protected SyncRowEnumerator(TRows rows)
        {
            Rows = rows;
        }

        /// <summary>The model for the row the cursor currently sits on, built on access.</summary>
        /// <remarks>
        /// Built here rather than cached by <see cref="MoveNext"/> because a class cannot hold a field
        /// of a <c>ref struct</c> <typeparamref name="T"/>. <c>foreach</c> reads this once per row, so
        /// the parse count is unchanged; reading it twice parses twice.
        /// </remarks>
        public T Current => Project();

        object? IEnumerator.Current => throw new NotSupportedException();

        /// <inheritdoc/>
        public bool MoveNext()
        {
            while (Rows.MoveNext())
            {
                switch (Classify())
                {
                    case ProjectionStep.Yield:
                        return true;
                    case ProjectionStep.Stop:
                        return false;
                }
            }
            return false;
        }

        /// <summary>Decides what the current row is (header, blank, data, end) without building a model.</summary>
        private protected abstract ProjectionStep Classify();

        /// <summary>Builds the model for a row <see cref="Classify"/> already resolved to <see cref="ProjectionStep.Yield"/>.</summary>
        private protected abstract T Project();

        /// <inheritdoc/>
        [SuppressMessage("Design", "CA1816:Dispose methods should call SuppressFinalize",
            Justification = "No finalizer exists on this type or any sealed derivative, so there is nothing to suppress.")]
        public void Dispose()
        {
            Rows.Dispose();
        }

        /// <inheritdoc/>
        public void Reset()
        {
            throw new NotSupportedException();
        }
    }
}
