using System.ComponentModel;
using System.Diagnostics.CodeAnalysis;
using ExcelReader.Core.Parser.Internal;
using ExcelReader.Core.Reader;

namespace ExcelReader.Core.Parser
{
    /// <summary>Base class supplying the shared row-advancement loop for an asynchronous, format-specific row enumerator.</summary>
    /// <typeparam name="T">The row model type each derived enumerator yields.</typeparam>
    /// <typeparam name="TSheet">The sheet type used to lazily open the row cursor.</typeparam>
    /// <typeparam name="TRows">The concrete row enumerator type this instance drives.</typeparam>
    /// <remarks>
    /// Mirrors <see cref="SyncRowEnumerator{T, TRows}"/>, plus lazy <typeparamref name="TRows"/>
    /// acquisition on the first <c>MoveNextAsync</c> and a sync-completion fast path: <c>MoveNextAsync</c> returns an already-completed <c>ValueTask</c>
    /// whenever the row-enumerator call and the projection both resolve synchronously (the common
    /// case), only falling to an awaiting continuation on a genuine buffer miss — this avoids paying
    /// for a second state machine on top of the row-enumerator's own (e.g.
    /// <c>XlsxWorkbook.Enumerator.MoveNextAsync</c> / <c>CsvReader.Enumerator.MoveNextAsync</c>).
    /// </remarks>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public abstract class AsyncRowEnumerator<T, TSheet, TRows> : IAsyncEnumerator<T>
        where T : allows ref struct
        where TSheet : IExcelSheet<TRows>
        where TRows : class, IExcelRowEnumerator
    {
        private readonly TSheet _sheet;
        private readonly CancellationToken _ct;
        private protected TRows? Rows;

        private protected AsyncRowEnumerator(TSheet sheet, CancellationToken ct)
        {
            _sheet = sheet;
            _ct = ct;
        }

        /// <summary>The model for the row the cursor currently sits on, built on access.</summary>
        /// <remarks>
        /// Built here rather than cached by <see cref="MoveNextAsync"/> because a class cannot hold a
        /// field of a <c>ref struct</c> <typeparamref name="T"/>, and such a value could not survive an
        /// <c>await</c> anyway. Building it outside the awaiting path is what lets a ref struct model
        /// flow through <see cref="IAsyncEnumerator{T}"/> at all.
        /// </remarks>
        public T Current => Project();

        /// <inheritdoc/>
        public ValueTask<bool> MoveNextAsync()
        {
            Rows ??= _sheet.GetAsyncEnumerator(_ct);
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
                switch (Classify())
                {
                    case ProjectionStep.Yield:
                        return new ValueTask<bool>(true);
                    case ProjectionStep.Stop:
                        return new ValueTask<bool>(false);
                }
            }
        }

        /// <summary>Decides what the current row is (header, blank, data, end) without building a model.</summary>
        private protected abstract ProjectionStep Classify();

        /// <summary>Builds the model for a row <see cref="Classify"/> already resolved to <see cref="ProjectionStep.Yield"/>.</summary>
        private protected abstract T Project();

        private async ValueTask<bool> AwaitThenContinueAsync(ValueTask<bool> pendingMoveNext)
        {
            if (!await pendingMoveNext.ConfigureAwait(false))
            {
                return false;
            }
            switch (Classify())
            {
                case ProjectionStep.Yield:
                    return true;
                case ProjectionStep.Stop:
                    return false;
            }
            return await MoveNextAsync().ConfigureAwait(false);
        }

        /// <inheritdoc/>
        [SuppressMessage("Design", "CA1816:Dispose methods should call SuppressFinalize",
            Justification = "No finalizer exists on this type or any sealed derivative, so there is nothing to suppress.")]
        public virtual ValueTask DisposeAsync()
        {
            return Rows is null ? ValueTask.CompletedTask : Rows.DisposeAsync();
        }
    }
}
