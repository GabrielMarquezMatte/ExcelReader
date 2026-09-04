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
    [SuppressMessage("Sonar", "S3881:IDisposable should be implemented correctly",
        Justification = "No unmanaged resources and no finalizer; every derived Enumerator is sealed and adds no disposal logic.")]
    public abstract class SyncRowEnumerator<T, TRows> : IEnumerator<T>
        where TRows : class, IExcelRowEnumerator
    {
        /// <summary>The underlying row cursor this enumerator advances.</summary>
        [SuppressMessage("Design", "CA1051:Do not declare visible instance fields",
            Justification = "Hot-path base class (MoveNext runs per row); a field avoids a property-call indirection in the tightest loop of the library.")]
        protected readonly TRows Rows;
        /// <summary>The most recently projected row model, returned by <see cref="Current"/>.</summary>
        [SuppressMessage("Design", "CA1051:Do not declare visible instance fields",
            Justification = "Hot-path base class (MoveNext runs per row); a field avoids a property-call indirection in the tightest loop of the library.")]
        protected T CurrentValue = default!;

        /// <summary>Initializes the base enumerator with the row cursor it will drive.</summary>
        /// <param name="rows">The row enumerator to advance and project from.</param>
        protected SyncRowEnumerator(TRows rows)
        {
            Rows = rows;
        }

        /// <inheritdoc/>
        public T Current => CurrentValue;

        object? IEnumerator.Current => CurrentValue;

        /// <inheritdoc/>
        public bool MoveNext()
        {
            while (Rows.MoveNext())
            {
                switch (Project())
                {
                    case ProjectionStep.Yield:
                        return true;
                    case ProjectionStep.Stop:
                        return false;
                }
            }
            return false;
        }

        private protected abstract ProjectionStep Project();

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
