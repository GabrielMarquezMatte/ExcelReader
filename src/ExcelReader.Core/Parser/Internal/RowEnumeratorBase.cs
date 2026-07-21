using System.Collections;
using System.Diagnostics.CodeAnalysis;
using ExcelReader.Core.Reader;

namespace ExcelReader.Core.Parser.Internal
{
    // Shared sync enumerator plumbing for every row-projecting IEnumerable<T> (Excel formats + CSV):
    // loop until Rows.MoveNext() is exhausted, project each row via the format-specific Project()
    // override, stopping early on ProjectionStep.Stop. Project() is the only thing that differs per
    // format (Excel walks Row.Cells generically via RowProjector<T>; CSV binds fields by dense index
    // via CsvRowProjector<T>) - see ExcelEnumerable<T,TReader,TEnumerator>.Enumerator and
    // CsvEnumerable<T>.Enumerator. Public because it is the base class of those public nested types
    // (a base class can never be less accessible than its derived type).
    [SuppressMessage("Design", "CA1034:Nested types should not be visible",
        Justification = "Base class of the public nested Enumerator types; not itself meant for direct external use.")]
    public abstract class SyncRowEnumerator<T, TRows> : IEnumerator<T>
        where TRows : class, IExcelRowEnumerator
    {
        [SuppressMessage("Performance", "HLQ011:ReadOnlyEnumeratorField",
            Justification = "TRows is constrained to `class` here, so it is always a reference type — no copy-on-mutate risk from a readonly field.")]
        protected readonly TRows Rows;
        protected T CurrentValue = default!;

        protected SyncRowEnumerator(TRows rows)
        {
            Rows = rows;
        }

        public T Current => CurrentValue;

        object? IEnumerator.Current => CurrentValue;

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

        [SuppressMessage("IDisposableAnalyzers.Correctness", "IDISP007:Don't dispose injected",
            Justification = "Rows is created for this enumerator alone by the enclosing enumerable's GetEnumerator (reader.GetEnumerator()) — owned here, not injected.")]
        public void Dispose()
        {
            Rows.Dispose();
        }

        public void Reset()
        {
            throw new NotSupportedException();
        }
    }

    // Shared async enumerator plumbing. Mirrors SyncRowEnumerator, plus the lazy TRows acquisition
    // (the reader's GetAsyncEnumeratorAsync may itself need to await) and the sync-completion fast
    // path: MoveNextAsync returns an already-completed ValueTask whenever the row-enumerator call and
    // the projection both resolve synchronously (the common case), only falling to an awaiting
    // continuation on a genuine buffer miss - this avoids paying for a second state machine on top of
    // the row-enumerator's own (e.g. XlsxReader.Enumerator.MoveNextAsync / CsvReader.Enumerator.MoveNextAsync).
    [SuppressMessage("Design", "CA1034:Nested types should not be visible",
        Justification = "Base class of the public nested AsyncEnumerator types; not itself meant for direct external use.")]
    public abstract class AsyncRowEnumerator<T, TReader, TRows> : IAsyncEnumerator<T>
        where TReader : IExcelRowReader<TRows>
        where TRows : class, IExcelRowEnumerator
    {
        // Borrowed: the caller owns the reader's lifetime. Only Rows (opened here) is disposed.
        [SuppressMessage("SharpSource", "SS066:Disposable field is not disposed", Justification = "Borrowed, not owned.")]
        private readonly TReader _reader;
        private readonly CancellationToken _ct;
        protected TRows? Rows;
        protected T CurrentValue = default!;

        protected AsyncRowEnumerator(TReader reader, CancellationToken ct)
        {
            _reader = reader;
            _ct = ct;
        }

        public T Current => CurrentValue;

        [SuppressMessage("SharpSource", "SS034:Use await to get the result of a Task",
            Justification = "The .Result access is guarded by IsCompletedSuccessfully immediately above it — never blocks.")]
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

        public ValueTask DisposeAsync()
        {
            return Rows is null ? ValueTask.CompletedTask : Rows.DisposeAsync();
        }
    }
}
