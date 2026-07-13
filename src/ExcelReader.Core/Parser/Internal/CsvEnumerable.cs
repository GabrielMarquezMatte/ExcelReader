using System.Collections;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using ExcelReader.Core.Enums;
using ExcelReader.Core.Reader;
using ExcelReader.Core.ValueObjects;

namespace ExcelReader.Core.Parser.Internal
{
    // CSV-specialized counterpart to ExcelEnumerable<T>. Because CSV rows are dense (columns 0..n-1,
    // no gaps) and every cell is text, this binds each property to a fixed field index at the header
    // row and then parses each data row in a single indexed pass — no Row.Cells enumeration, no
    // CellDesc re-walk, no per-cell binding search. It reuses the same compiled ColumnParser<T>
    // delegates as the generic parser, but from the CSV type-map (dates parse text, not serials).
    [SuppressMessage("Design", "CA1034:Nested types should not be visible",
        Justification = "Public nested struct Enumerator is the standard foreach pattern.")]
    public sealed class CsvEnumerable<T> : IEnumerable<T>, IAsyncEnumerable<T>
    {
        private readonly CsvReader _reader;
        private readonly ExcelParserConfig _config;
        private readonly CancellationToken _ct;

        internal CsvEnumerable(CsvReader reader, ExcelParserConfig config, CancellationToken ct = default)
        {
            _reader = reader;
            _config = config;
            _ct = ct;
        }

        [SuppressMessage("IDisposableAnalyzers.Correctness", "IDISP015:Member should not return created and cached instance",
            Justification = "Each call creates a fresh enumerator; no caching.")]
        public Enumerator GetEnumerator()
        {
            TypeMapInfo<T> info = TypeMapper<T>.GetCsvInfo();
            CsvReader.Enumerator rows = _reader.GetEnumerator();
            return new Enumerator(rows, info, _config.ColumnNameComparer, _config.HeaderNormalization, _config.HeaderRow, _config.Culture);
        }

        IEnumerator<T> IEnumerable<T>.GetEnumerator()
        {
            return GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        [SuppressMessage("Performance", "HLQ006:GetAsyncEnumerator should return a value type",
            Justification = "Async enumerator requires a class to host the async state machine.")]
        IAsyncEnumerator<T> IAsyncEnumerable<T>.GetAsyncEnumerator(CancellationToken cancellationToken)
        {
            return GetAsyncEnumerator(cancellationToken);
        }

        public AsyncEnumerator GetAsyncEnumerator(CancellationToken cancellationToken = default)
        {
            TypeMapInfo<T> info = TypeMapper<T>.GetCsvInfo();
            CancellationToken effective = cancellationToken.CanBeCanceled ? cancellationToken : _ct;
            return new AsyncEnumerator(_reader, info, _config.ColumnNameComparer, _config.HeaderNormalization, _config.HeaderRow, _config.Culture, effective);
        }

        public struct Enumerator : IEnumerator<T>
        {
            [SuppressMessage("IDisposableAnalyzers.Correctness", "IDISP006:Implement IDisposable",
                Justification = "Struct implements IDisposable; rows disposed in Dispose().")]
            private readonly CsvReader.Enumerator _rows;
            private CsvRowProjector<T> _projector;
            private T _current = default!;

            internal Enumerator(
                CsvReader.Enumerator rows,
                TypeMapInfo<T> typeInfo,
                StringComparer comparer,
                HeaderNormalization normalization,
                int headerRow,
                IFormatProvider provider)
            {
                _rows = rows;
                _projector = new CsvRowProjector<T>(typeInfo, comparer, normalization, headerRow, provider);
            }

            public readonly T Current => _current;

            readonly object? IEnumerator.Current => _current;

            public bool MoveNext()
            {
                while (_rows.MoveNext())
                {
                    switch (_projector.Advance(_rows, ref _current))
                    {
                        case ProjectionStep.Yield:
                            return true;
                        case ProjectionStep.Stop:
                            return false;
                    }
                }
                return false;
            }

            [SuppressMessage("IDisposableAnalyzers.Correctness", "IDISP007:Don't dispose injected",
                Justification = "The enumerator owns _rows; it was created by GetEnumerator, not injected from outside.")]
            public readonly void Dispose()
            {
                _rows.Dispose();
            }

            public void Reset()
            {
                throw new NotSupportedException();
            }
        }

        public sealed class AsyncEnumerator : IAsyncEnumerator<T>
        {
            // Borrowed: the caller owns the CsvReader's lifetime. Only _rows (opened here) is disposed.
            [SuppressMessage("SharpSource", "SS066:Disposable field is not disposed", Justification = "Borrowed, not owned.")]
            private readonly CsvReader _reader;
            private readonly CancellationToken _ct;
            private CsvRowProjector<T> _projector;
            private CsvReader.Enumerator? _rows;
            private T _current = default!;

            internal AsyncEnumerator(
                CsvReader reader,
                TypeMapInfo<T> typeInfo,
                StringComparer comparer,
                HeaderNormalization normalization,
                int headerRow,
                IFormatProvider provider,
                CancellationToken ct)
            {
                _reader = reader;
                _projector = new CsvRowProjector<T>(typeInfo, comparer, normalization, headerRow, provider);
                _ct = ct;
            }

            public T Current => _current;

            // Non-async fast path once _rows exists — see ExcelEnumerable.AsyncEnumerator.MoveNextAsync
            // for the rationale: CsvReader.Enumerator.MoveNextAsync already resolves synchronously for
            // ~99.9% of records, so this avoids paying for a second state machine on top of that one.
            [SuppressMessage("SharpSource", "SS034:Use await to get the result of a Task",
                Justification = "The .Result access is guarded by IsCompletedSuccessfully immediately above it — never blocks.")]
            [SuppressMessage("VisualStudio.Threading", "VSTHRD103:Result synchronously blocks",
                Justification = "The .Result access is guarded by IsCompletedSuccessfully immediately above it — never blocks.")]
            public ValueTask<bool> MoveNextAsync()
            {
                if (_rows is null)
                {
                    return AdvanceAsync();
                }
                while (true)
                {
                    ValueTask<bool> moveTask = _rows.MoveNextAsync();
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
                _rows = await _reader.GetAsyncEnumeratorAsync(_ct).ConfigureAwait(false);
                return await MoveNextAsync().ConfigureAwait(false);
            }

            // Synchronous projection step: the ref-struct Cells never escape this call, so no span
            // is held across the await in AdvanceAsync.
            private ProjectionStep Project()
            {
                return _projector.Advance(_rows!, ref _current);
            }

            public ValueTask DisposeAsync()
            {
                return _rows is null ? ValueTask.CompletedTask : _rows.DisposeAsync();
            }
        }
    }

    // The per-row state machine for CSV: skip rows before the header, bind property -> field index at
    // the header row, then project each data row by direct indexed field access.
    internal struct CsvRowProjector<T>
    {
        private readonly TypeMapInfo<T> _typeInfo;
        private readonly StringComparer _comparer;
        private readonly HeaderNormalization _normalization;
        private readonly int _headerRow;
        private readonly IFormatProvider _provider;
        // fieldParsers[i] is the parser bound to field i, or null if field i is unmapped.
        private ColumnParser<T>[]? _fieldParsers;
        // Fields whose bound property is [ExcelRequired] without AllowEmpty; each data row must carry
        // a non-empty value there. Empty when no bound column requires a value.
        private (int Field, string Name)[] _requiredFields;
        private int _rowNumber;

        internal CsvRowProjector(TypeMapInfo<T> typeInfo, StringComparer comparer, HeaderNormalization normalization, int headerRow, IFormatProvider provider)
        {
            _typeInfo = typeInfo;
            _comparer = comparer;
            _normalization = normalization;
            _headerRow = headerRow;
            _provider = provider;
            _requiredFields = [];
        }

        internal ProjectionStep Advance(CsvReader.Enumerator rows, ref T model)
        {
            _rowNumber++;
            if (_rowNumber < _headerRow)
            {
                return ProjectionStep.Skip;
            }
            if (_rowNumber == _headerRow)
            {
                BuildColumnMap(rows);
                return ProjectionStep.Skip;
            }
            if (_fieldParsers is null)
            {
                return ProjectionStep.Stop;
            }
            // The raw CSV reader intentionally exposes a terminal blank line as one empty field.
            // Typed projection treats it as absent so it cannot yield a phantom model or fail Required.
            if (rows.FieldCount == 1 && rows.FieldAt(0).Type == CellType.Empty)
            {
                return ProjectionStep.Skip;
            }
            model = _typeInfo.CreateInstance();
            ParseCurrentRow(rows, ref model);
            return ProjectionStep.Yield;
        }

        private void BuildColumnMap(CsvReader.Enumerator rows)
        {
            int fieldCount = rows.FieldCount;
            int propertyCount = _typeInfo.PropertyCount;
            var parsers = new ColumnParser<T>[fieldCount];
            int[] aliasByProp = new int[propertyCount];
            int[] fieldByProp = new int[propertyCount];
            Array.Fill(aliasByProp, int.MaxValue);
            Array.Fill(fieldByProp, -1);

            for (int i = 0; i < fieldCount; i++)
            {
                Cell cell = rows.FieldAt(i);
                string header = _normalization.Apply(cell.GetString());
                if (string.IsNullOrEmpty(header))
                {
                    continue;
                }
                if (!_typeInfo.TryFindHeader(header, _comparer, _normalization, out HeaderMatch<T> match))
                {
                    continue;
                }
                if (match.AliasIndex >= aliasByProp[match.PropertyIndex])
                {
                    continue;
                }
                // A lower-priority alias already bound this property to another field: unbind it so
                // each property maps to exactly one field (mirrors the generic RowProjector).
                if (fieldByProp[match.PropertyIndex] >= 0)
                {
                    parsers[fieldByProp[match.PropertyIndex]] = null!;
                }
                aliasByProp[match.PropertyIndex] = match.AliasIndex;
                fieldByProp[match.PropertyIndex] = i;
                parsers[i] = match.Parser;
            }

            _typeInfo.ValidateRequiredColumns(aliasByProp);

            List<(int, string)>? required = null;
            for (int p = 0; p < propertyCount; p++)
            {
                if (fieldByProp[p] >= 0 && _typeInfo.RequiresValue(p))
                {
                    (required ??= []).Add((fieldByProp[p], _typeInfo.DisplayName(p)));
                }
            }

            _fieldParsers = parsers;
            _requiredFields = required is null ? [] : [.. required];
        }

        private readonly void ParseCurrentRow(CsvReader.Enumerator rows, ref T model)
        {
            ColumnParser<T>[] parsers = _fieldParsers!;
            int fieldCount = rows.FieldCount;
            int limit = Math.Min(fieldCount, parsers.Length);
            for (int i = 0; i < limit; i++)
            {
                ColumnParser<T>? parser = parsers[i];
                if (parser is null)
                {
                    continue;
                }
                Cell cell = rows.FieldAt(i);
                if (cell.Type != CellType.Empty)
                {
                    parser(ref model, in cell, false, _provider);
                }
            }

            foreach ((int field, string name) in _requiredFields)
            {
                if (field >= fieldCount || rows.FieldAt(field).Type == CellType.Empty)
                {
                    throw new InvalidOperationException(
                        $"Required column '{name}' has no value in row {_rowNumber.ToString(CultureInfo.InvariantCulture)}.");
                }
            }
        }
    }
}
