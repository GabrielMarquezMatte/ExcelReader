using System.Collections;
using System.Diagnostics.CodeAnalysis;
using ExcelReader.Core.Enums;
using ExcelReader.Core.Reader;
using ExcelReader.Core.ValueObjects;

namespace ExcelReader.Core.Parser.Internal
{
    /// <summary>Lazily projects CSV rows into <typeparamref name="T"/> instances by binding each property to a fixed field index, for both synchronous and asynchronous enumeration.</summary>
    /// <typeparam name="T">The row model type to bind each CSV row to.</typeparam>
    /// <remarks>
    /// The CSV-specialized counterpart to <see cref="ExcelEnumerable{T}"/>. Because CSV rows are dense
    /// (columns 0..n-1, no gaps) and every cell is text, each data row is parsed in a single indexed
    /// pass — no <c>Row.Cells</c> enumeration, no cell-descriptor re-walk, no per-cell binding search. It
    /// reuses the same compiled <c>ColumnParser&lt;T&gt;</c> delegates as the generic parser, but from the
    /// CSV type-map, where dates parse text rather than Excel serial numbers.
    /// </remarks>
    [SuppressMessage("Design", "CA1034:Nested types should not be visible",
        Justification = "Public nested Enumerator/AsyncEnumerator are the standard foreach/await-foreach pattern.")]
    public sealed class CsvEnumerable<T> : IEnumerable<T>, IAsyncEnumerable<T>
    {
        private readonly CsvReader _reader;
        private readonly ExcelParserConfig _config;
        private readonly CancellationToken _ct;
        // Resolved once here, in the constructor — never in GetEnumerator()/GetAsyncEnumerator() — see
        // ExcelEnumerable<T,TReader,TEnumerator>'s matching field for why: a trimmer/AOT analyzer decides
        // reachability per method, not per field, so TypeMapper<T>.GetCsvInfo()'s reflection must live
        // only in the constructor ExcelParser<T> uses, never in a method the AOT-clean ExcelMappedParser<T>
        // path also calls. The mapped path also uses one map for every reader, so unlike the reflection
        // path's dedicated CSV map (TypeMapper<T>.GetCsvInfo()), a CSV date property here reads with
        // whatever reader its ConfigureExcelRowMap chose — a caller who wants text dates from CSV picks
        // ExcelCellReaders.DateTimeText explicitly.
        private readonly TypeMapInfo<T> _info;

        [RequiresUnreferencedCode("Typed parsing reflects over T's public properties, which trimming may remove.")]
        [RequiresDynamicCode("Typed parsing binds property setters at runtime (MethodInfo.CreateDelegate / MakeGenericMethod).")]
        internal CsvEnumerable(CsvReader reader, ExcelParserConfig config, CancellationToken ct = default)
        {
            _reader = reader;
            _config = config;
            _info = TypeMapper<T>.GetCsvInfo();
            _ct = ct;
        }

        internal CsvEnumerable(CsvReader reader, ExcelParserConfig config, TypeMapInfo<T> explicitInfo, CancellationToken ct = default)
        {
            _reader = reader;
            _config = config;
            _info = explicitInfo;
            _ct = ct;
        }

        /// <inheritdoc cref="IEnumerable{T}.GetEnumerator"/>
        [SuppressMessage("Performance", "HLQ006:GetEnumerator should return a value type",
            Justification = "Enumerator is a class so the sync and async paths can share the SyncRowEnumerator/AsyncRowEnumerator base plumbing.")]
        [SuppressMessage("IDisposableAnalyzers.Correctness", "IDISP001:Dispose created",
            Justification = "rows is handed to the returned Enumerator, which owns and disposes it.")]
        public Enumerator GetEnumerator()
        {
            CsvReader.Enumerator rows = _reader.GetEnumerator();
            return new Enumerator(rows, _info, _config.ColumnNameComparer, _config.HeaderNormalization, _config.HeaderRow, _config.Culture, _config.ThrowOnParseFailure);
        }

        IEnumerator<T> IEnumerable<T>.GetEnumerator()
        {
            return GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        IAsyncEnumerator<T> IAsyncEnumerable<T>.GetAsyncEnumerator(CancellationToken cancellationToken)
        {
            return GetAsyncEnumerator(cancellationToken);
        }

        /// <inheritdoc cref="IAsyncEnumerable{T}.GetAsyncEnumerator"/>
        [SuppressMessage("Performance", "HLQ006:GetAsyncEnumerator should return a value type",
            Justification = "Async enumerator requires a class to host the async state machine.")]
        public AsyncEnumerator GetAsyncEnumerator(CancellationToken cancellationToken = default)
        {
            CancellationToken effective = cancellationToken.CanBeCanceled ? cancellationToken : _ct;
            return new AsyncEnumerator(_reader, _info, _config.ColumnNameComparer, _config.HeaderNormalization, _config.HeaderRow, _config.Culture, _config.ThrowOnParseFailure, effective);
        }

        /// <summary>Enumerates CSV rows synchronously, projecting each into a <typeparamref name="T"/> instance by fixed field index.</summary>
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

        /// <summary>Enumerates CSV rows asynchronously, projecting each into a <typeparamref name="T"/> instance by fixed field index.</summary>
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

    // The header-derived, per-field binding state of a CSV projection, lifted out of
    // CsvRowProjector<T> so the parallel path can bind a header once and share one instance across
    // every worker. Every array here is written during construction and only read afterwards, which
    // is what makes sharing it across threads safe without a lock.
    internal sealed class CsvBoundColumnMap<T>
    {
        internal CsvBoundColumnMap(
            ColumnParser<T>[] fieldParsers,
            string?[] fieldNames,
            bool[] fieldRequired,
            (int Field, string Name)[] requiredFields)
        {
            FieldParsers = fieldParsers;
            FieldNames = fieldNames;
            FieldRequired = fieldRequired;
            RequiredFields = requiredFields;
        }

        internal ColumnParser<T>[] FieldParsers { get; }

        internal string?[] FieldNames { get; }

        internal bool[] FieldRequired { get; }

        internal (int Field, string Name)[] RequiredFields { get; }
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

        // Parallel-path constructor: the header was already read and bound once by CsvHeaderBinder,
        // so there is no header row to find and no rows to skip — every record this projector sees
        // is data. _headerRow is set to -1 to make that explicit; the Advance fast path below keys
        // off _fieldParsers already being non-null, exactly as the index-based path does.
        internal CsvRowProjector(TypeMapInfo<T> typeInfo, CsvBoundColumnMap<T> map, IFormatProvider provider, bool throwOnParseFailure)
        {
            _typeInfo = typeInfo;
            _comparer = StringComparer.Ordinal;
            _normalization = HeaderNormalization.None;
            _headerRow = -1;
            _provider = provider;
            _throwOnParseFailure = throwOnParseFailure;
            _fieldParsers = map.FieldParsers;
            _fieldNames = map.FieldNames;
            _fieldRequired = map.FieldRequired;
            _requiredFields = map.RequiredFields;
            _rowNumber = 0;
        }

        internal ProjectionStep Advance(CsvReader.Enumerator rows, ref T model)
        {
            if (_headerRow < 0)
            {
                _rowNumber++;
                if (rows.FieldCount == 1 && rows.FieldAt(0).Type == CellType.Empty)
                {
                    return ProjectionStep.Skip;
                }
                model = _typeInfo.CreateInstance();
                ParseCurrentRow(rows, ref model);
                return ProjectionStep.Yield;
            }
            if (_typeInfo.IsIndexBased)
            {
                if (_fieldParsers is null)
                {
                    BuildIndexColumnMap();
                }
                _rowNumber++;
                // Same terminal-blank-line treatment as the header path below: a trailing empty line
                // must not yield a phantom model or trip a required-column check.
                if (rows.FieldCount == 1 && rows.FieldAt(0).Type == CellType.Empty)
                {
                    return ProjectionStep.Skip;
                }
                model = _typeInfo.CreateInstance();
                ParseCurrentRow(rows, ref model);
                return ProjectionStep.Yield;
            }
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
            CsvBoundColumnMap<T> map = BuildBoundMap(rows, _typeInfo, _comparer, _normalization);
            _fieldParsers = map.FieldParsers;
            _fieldNames = map.FieldNames;
            _fieldRequired = map.FieldRequired;
            _requiredFields = map.RequiredFields;
        }

        // Binds the current header row to T's properties, by field index. Shared by the sequential
        // path (BuildColumnMap, above) and CsvHeaderBinder's parallel path — one implementation of
        // header binding, reused rather than duplicated.
        internal static CsvBoundColumnMap<T> BuildBoundMap(CsvReader.Enumerator rows, TypeMapInfo<T> typeInfo, StringComparer comparer, HeaderNormalization normalization)
        {
            int fieldCount = rows.FieldCount;
            int propertyCount = typeInfo.PropertyCount;
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
                string header = normalization.Apply(cell.GetString());
                if (string.IsNullOrEmpty(header))
                {
                    continue;
                }
                if (!typeInfo.TryFindHeader(header, comparer, normalization, out HeaderMatch<T> match))
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
                names[i] = typeInfo.DisplayName(match.PropertyIndex);
                required[i] = typeInfo.RequiresValue(match.PropertyIndex);
            }

            typeInfo.ValidateRequiredColumns(aliasByProp);

            List<(int, string)>? requiredFields = null;
            for (int p = 0; p < propertyCount; p++)
            {
                if (fieldByProp[p] >= 0 && typeInfo.RequiresValue(p))
                {
                    (requiredFields ??= []).Add((fieldByProp[p], typeInfo.DisplayName(p)));
                }
            }

            return new CsvBoundColumnMap<T>(parsers, names, required, requiredFields is null ? [] : [.. requiredFields]);
        }

        private void BuildIndexColumnMap()
        {
            ColumnBinding<T>[] bindings = _typeInfo.IndexBindings;
            int width = bindings.Length == 0 ? 0 : bindings[^1].Column + 1;
            var parsers = new ColumnParser<T>[width];
            var names = new string?[width];
            var required = new bool[width];
            List<(int, string)>? requiredFields = null;
            foreach (ColumnBinding<T> binding in bindings)
            {
                parsers[binding.Column] = binding.Parser;
                names[binding.Column] = binding.Name;
                required[binding.Column] = binding.RequireValue;
                if (binding.RequireValue)
                {
                    (requiredFields ??= []).Add((binding.Column, binding.Name));
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
