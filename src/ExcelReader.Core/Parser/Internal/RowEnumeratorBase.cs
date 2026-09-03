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
    /// accessible than its derived type).
    /// </remarks>
    [SuppressMessage("Design", "CA1063:Implement IDisposable correctly",
        Justification = "No unmanaged resources and no finalizer; every derived Enumerator is sealed and adds no disposal logic, so the full Dispose(bool) pattern buys nothing here.")]
    [SuppressMessage("IDisposableAnalyzers.Correctness", "IDISP025:Class with no virtual dispose method should be sealed",
        Justification = "Abstract by design (base of the public nested Enumerator types); every concrete derivative is itself sealed.")]
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

    /// <summary>Base class supplying the shared row-advancement loop for an asynchronous, format-specific row enumerator.</summary>
    /// <typeparam name="T">The row model type each derived enumerator yields.</typeparam>
    /// <typeparam name="TReader">The concrete row reader type used to lazily open the row cursor.</typeparam>
    /// <typeparam name="TRows">The concrete row enumerator type this instance drives.</typeparam>
    /// <remarks>
    /// Mirrors <see cref="SyncRowEnumerator{T, TRows}"/>, plus lazy <typeparamref name="TRows"/>
    /// acquisition (the reader's <c>GetAsyncEnumeratorAsync</c> may itself need to await) and a
    /// sync-completion fast path: <c>MoveNextAsync</c> returns an already-completed <c>ValueTask</c>
    /// whenever the row-enumerator call and the projection both resolve synchronously (the common
    /// case), only falling to an awaiting continuation on a genuine buffer miss — this avoids paying
    /// for a second state machine on top of the row-enumerator's own (e.g.
    /// <c>XlsxReader.Enumerator.MoveNextAsync</c> / <c>CsvReader.Enumerator.MoveNextAsync</c>).
    /// </remarks>
    [SuppressMessage("IDisposableAnalyzers.Correctness", "IDISP026:Class with no virtual DisposeAsyncCore method should be sealed",
        Justification = "Abstract by design (base of the public nested AsyncEnumerator types); every concrete derivative is itself sealed.")]
    public abstract class AsyncRowEnumerator<T, TReader, TRows> : IAsyncEnumerator<T>
        where TReader : IExcelRowReader<TRows>
        where TRows : class, IExcelRowEnumerator
    {
        // Borrowed: the caller owns the reader's lifetime. Only Rows (opened here) is disposed.
        private readonly TReader _reader;
        private readonly CancellationToken _ct;
        /// <summary>The underlying row cursor this enumerator advances, opened lazily on the first call to <see cref="MoveNextAsync"/>.</summary>
        [SuppressMessage("Design", "CA1051:Do not declare visible instance fields",
            Justification = "Hot-path base class (MoveNextAsync runs per row); a field avoids a property-call indirection in the tightest loop of the library.")]
        protected TRows? Rows;
        /// <summary>The most recently projected row model, returned by <see cref="Current"/>.</summary>
        [SuppressMessage("Design", "CA1051:Do not declare visible instance fields",
            Justification = "Hot-path base class (MoveNextAsync runs per row); a field avoids a property-call indirection in the tightest loop of the library.")]
        protected T CurrentValue = default!;

        /// <summary>Initializes the base enumerator with the reader it will lazily open a row cursor from.</summary>
        /// <param name="reader">The row reader used to open the row cursor on first advancement.</param>
        /// <param name="ct">The cancellation token to use when opening the row cursor, if the caller does not supply one to <see cref="MoveNextAsync"/>.</param>
        protected AsyncRowEnumerator(TReader reader, CancellationToken ct)
        {
            _reader = reader;
            _ct = ct;
        }

        /// <inheritdoc/>
        public T Current => CurrentValue;

        /// <inheritdoc/>
        [SuppressMessage("VisualStudio.Threading", "VSTHRD103:Result synchronously blocks",
            Justification = "The .Result access is guarded by IsCompletedSuccessfully immediately above it — never blocks.")]
        public ValueTask<bool> MoveNextAsync()
        {
            if (Rows is null)
            {
                return AdvanceAsync();
            }
            while (true)
            {
                ValueTask<bool> moveTask = Rows.MoveNextAsync();
                if (!moveTask.IsCompletedSuccessfully)
                {
                    return AwaitThenContinueAsync(moveTask);
                }
                if (!moveTask.Result)
                {
                    return new ValueTask<bool>(false);
                }
                switch (Project())
                {
                    case ProjectionStep.Yield:
                        return new ValueTask<bool>(true);
                    case ProjectionStep.Stop:
                        return new ValueTask<bool>(false);
                        // Skip: loop again, still synchronous.
                }
            }
        }

        private protected abstract ProjectionStep Project();

        private async ValueTask<bool> AwaitThenContinueAsync(ValueTask<bool> pendingMoveNext)
        {
            if (!await pendingMoveNext.ConfigureAwait(false))
            {
                return false;
            }
            switch (Project())
            {
                case ProjectionStep.Yield:
                    return true;
                case ProjectionStep.Stop:
                    return false;
            }
            return await MoveNextAsync().ConfigureAwait(false); // Skip: resume the fast path.
        }

        private async ValueTask<bool> AdvanceAsync()
        {
            Rows = await _reader.GetAsyncEnumeratorAsync(_ct).ConfigureAwait(false);
            return await MoveNextAsync().ConfigureAwait(false);
        }

        /// <inheritdoc/>
        [SuppressMessage("Design", "CA1816:Dispose methods should call SuppressFinalize",
            Justification = "No finalizer exists on this type or any sealed derivative, so there is nothing to suppress.")]
        // Virtual so an enumerator that also owns the reader it was handed can close it here, instead
        // of the caller wrapping the whole enumerable in a second async iterator to get an
        // `await using` — that wrapper re-yields every row through another state machine, which costs
        // more per row than the disposal it exists to perform (see ParallelCsvFactory.Sequential).
        public virtual ValueTask DisposeAsync()
        {
            return Rows is null ? ValueTask.CompletedTask : Rows.DisposeAsync();
        }
    }
}
