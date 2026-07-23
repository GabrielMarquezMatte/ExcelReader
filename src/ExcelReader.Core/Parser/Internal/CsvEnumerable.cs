using System.Collections;
using System.Diagnostics.CodeAnalysis;
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
        Justification = "Public nested Enumerator/AsyncEnumerator are the standard foreach/await-foreach pattern.")]
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
        [SuppressMessage("Performance", "HLQ006:GetEnumerator should return a value type",
            Justification = "Enumerator is a class so the sync and async paths can share the SyncRowEnumerator/AsyncRowEnumerator base plumbing.")]
        [SuppressMessage("IDisposableAnalyzers.Correctness", "IDISP001:Dispose created",
            Justification = "rows is handed to the returned Enumerator, which owns and disposes it.")]
        public Enumerator GetEnumerator()
        {
            TypeMapInfo<T> info = TypeMapper<T>.GetCsvInfo();
            CsvReader.Enumerator rows = _reader.GetEnumerator();
            return new Enumerator(rows, info, _config.ColumnNameComparer, _config.HeaderNormalization, _config.HeaderRow, _config.Culture, _config.ThrowOnParseFailure);
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

        [SuppressMessage("Performance", "HLQ006:GetAsyncEnumerator should return a value type",
            Justification = "Async enumerator requires a class to host the async state machine.")]
        public AsyncEnumerator GetAsyncEnumerator(CancellationToken cancellationToken = default)
        {
            TypeMapInfo<T> info = TypeMapper<T>.GetCsvInfo();
            CancellationToken effective = cancellationToken.CanBeCanceled ? cancellationToken : _ct;
            return new AsyncEnumerator(_reader, info, _config.ColumnNameComparer, _config.HeaderNormalization, _config.HeaderRow, _config.Culture, _config.ThrowOnParseFailure, effective);
        }

        public sealed class Enumerator : SyncRowEnumerator<T, CsvReader.Enumerator>
        {
            private CsvRowProjector<T> _projector;

            internal Enumerator(
                CsvReader.Enumerator rows,
                TypeMapInfo<T> typeInfo,
                StringComparer comparer,
                HeaderNormalization normalization,
                int headerRow,
                IFormatProvider provider,
                bool throwOnParseFailure = false)
                : base(rows)
            {
                _projector = new CsvRowProjector<T>(typeInfo, comparer, normalization, headerRow, provider, throwOnParseFailure);
            }

            private protected override ProjectionStep Project()
            {
                return _projector.Advance(Rows, ref CurrentValue);
            }
        }

        public sealed class AsyncEnumerator : AsyncRowEnumerator<T, CsvReader, CsvReader.Enumerator>
        {
            private CsvRowProjector<T> _projector;

            internal AsyncEnumerator(
                CsvReader reader,
                TypeMapInfo<T> typeInfo,
                StringComparer comparer,
                HeaderNormalization normalization,
                int headerRow,
                IFormatProvider provider,
                bool throwOnParseFailure,
                CancellationToken ct)
                : base(reader, ct)
            {
                _projector = new CsvRowProjector<T>(typeInfo, comparer, normalization, headerRow, provider, throwOnParseFailure);
            }

            // Synchronous projection step: the ref-struct Cells never escape this call, so no span
            // is held across the await in the base class's AdvanceAsync.
            private protected override ProjectionStep Project()
            {
                return _projector.Advance(Rows!, ref CurrentValue);
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
        private readonly bool _throwOnParseFailure;
        // fieldParsers[i] is the parser bound to field i, or null if field i is unmapped.
        private ColumnParser<T>[]? _fieldParsers;
        // Parallel to _fieldParsers: the bound property's display name (for ExcelParseException) and
        // whether it is [ExcelRequired] without AllowEmpty, indexed by field i. Only meaningful where
        // _fieldParsers[i] is non-null.
        private string?[] _fieldNames;
        private bool[] _fieldRequired;
        // Fields whose bound property is [ExcelRequired] without AllowEmpty; each data row must carry
        // a non-empty value there. Empty when no bound column requires a value.
        private (int Field, string Name)[] _requiredFields;
        private int _rowNumber;

        internal CsvRowProjector(TypeMapInfo<T> typeInfo, StringComparer comparer, HeaderNormalization normalization, int headerRow, IFormatProvider provider, bool throwOnParseFailure = false)
        {
            _typeInfo = typeInfo;
            _comparer = comparer;
            _normalization = normalization;
            _headerRow = headerRow;
            _provider = provider;
            _throwOnParseFailure = throwOnParseFailure;
            _fieldNames = [];
            _fieldRequired = [];
            _requiredFields = [];
        }

        internal ProjectionStep Advance(CsvReader.Enumerator rows, ref T model)
        {
            ProjectionStep step = ProjectionRules.ClassifyRow(ref _rowNumber, _headerRow, _fieldParsers is not null);
            if (step == ProjectionStep.BuildMap)
            {
                BuildColumnMap(rows);
                return ProjectionStep.Skip;
            }
            if (step != ProjectionStep.Yield)
            {
                return step;
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
            var names = new string?[fieldCount];
            var required = new bool[fieldCount];
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
                    int previousField = fieldByProp[match.PropertyIndex];
                    parsers[previousField] = null!;
                    names[previousField] = null;
                    required[previousField] = false;
                }
                aliasByProp[match.PropertyIndex] = match.AliasIndex;
                fieldByProp[match.PropertyIndex] = i;
                parsers[i] = match.Parser;
                names[i] = _typeInfo.DisplayName(match.PropertyIndex);
                required[i] = _typeInfo.RequiresValue(match.PropertyIndex);
            }

            _typeInfo.ValidateRequiredColumns(aliasByProp);

            List<(int, string)>? requiredFields = null;
            for (int p = 0; p < propertyCount; p++)
            {
                if (fieldByProp[p] >= 0 && _typeInfo.RequiresValue(p))
                {
                    (requiredFields ??= []).Add((fieldByProp[p], _typeInfo.DisplayName(p)));
                }
            }

            _fieldParsers = parsers;
            _fieldNames = names;
            _fieldRequired = required;
            _requiredFields = requiredFields is null ? [] : [.. requiredFields];
        }

        // On a parse failure (non-empty cell, parser returned false): throws ExcelParseException when
        // _throwOnParseFailure is set; otherwise, a required field fails the same "missing required
        // value" way a blank one does (F3) instead of silently keeping the model's default.
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
                if (cell.Type == CellType.Empty || parser(ref model, in cell, false, _provider))
                {
                    continue;
                }
                if (_throwOnParseFailure)
                {
                    throw new ExcelParseException(_rowNumber, _fieldNames[i]!, cell.GetString());
                }
                if (_fieldRequired[i])
                {
                    throw ProjectionRules.MissingRequiredValue(_fieldNames[i]!, _rowNumber);
                }
            }

            foreach ((int field, string name) in _requiredFields)
            {
                if (field >= fieldCount || rows.FieldAt(field).Type == CellType.Empty)
                {
                    throw ProjectionRules.MissingRequiredValue(name, _rowNumber);
                }
            }
        }
    }
}
