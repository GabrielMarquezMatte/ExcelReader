using System.Collections;
using System.Diagnostics.CodeAnalysis;
using ExcelReader.Core.Enums;
using ExcelReader.Core.Parser.Internal;
using ExcelReader.Core.Reader;
using ExcelReader.Core.ValueObjects;

namespace ExcelReader.Core.Parser
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
        where T : allows ref struct
    {
        private readonly CsvReader _reader;
        private readonly ExcelParserConfig _config;
        private readonly CancellationToken _ct;
        private readonly TypeMapInfo<T> _info;
        private readonly bool _ownsReader;

        [RequiresUnreferencedCode("Typed parsing reflects over T's public properties, which trimming may remove.")]
        [RequiresDynamicCode("Typed parsing binds property setters at runtime (MethodInfo.CreateDelegate / MakeGenericMethod).")]
        internal CsvEnumerable(CsvReader reader, ExcelParserConfig config, bool ownsReader, CancellationToken ct)
        {
            _reader = reader;
            _config = config;
            _info = TypeMapper<T>.GetCsvInfo();
            _ct = ct;
            _ownsReader = ownsReader;
        }

        internal CsvEnumerable(CsvReader reader, ExcelParserConfig config, TypeMapInfo<T> explicitInfo)
        {
            _reader = reader;
            _config = config;
            _info = explicitInfo;
        }

        /// <inheritdoc cref="IEnumerable{T}.GetEnumerator"/>
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
        public AsyncEnumerator GetAsyncEnumerator(CancellationToken cancellationToken = default)
        {
            CancellationToken effective = cancellationToken.CanBeCanceled ? cancellationToken : _ct;
            return new AsyncEnumerator(_reader, _info, _config.ColumnNameComparer, _config.HeaderNormalization, _config.HeaderRow, _config.Culture, _config.ThrowOnParseFailure, _ownsReader, effective);
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

            private protected override ProjectionStep Classify()
            {
                return _projector.Classify(Rows);
            }

            private protected override T Project()
            {
                return _projector.Project(Rows);
            }
        }

        /// <summary>Enumerates CSV rows asynchronously, projecting each into a <typeparamref name="T"/> instance by fixed field index.</summary>
        public sealed class AsyncEnumerator : AsyncRowEnumerator<T, CsvReader, CsvReader.Enumerator>
        {
            private CsvRowProjector<T> _projector;
            private readonly CsvReader? _ownedReader;

            internal AsyncEnumerator(
                CsvReader reader,
                TypeMapInfo<T> typeInfo,
                StringComparer comparer,
                HeaderNormalization normalization,
                int headerRow,
                IFormatProvider provider,
                bool throwOnParseFailure,
                bool ownsReader,
                CancellationToken ct)
                : base(reader, ct)
            {
                _projector = new CsvRowProjector<T>(typeInfo, comparer, normalization, headerRow, provider, throwOnParseFailure);
                _ownedReader = ownsReader ? reader : null;
            }

            /// <inheritdoc/>
            public override async ValueTask DisposeAsync()
            {
                await base.DisposeAsync().ConfigureAwait(false);
                if (_ownedReader is not null)
                {
                    await _ownedReader.DisposeAsync().ConfigureAwait(false);
                }
            }

            private protected override ProjectionStep Classify()
            {
                return _projector.Classify(Rows!);
            }

            private protected override T Project()
            {
                return _projector.Project(Rows!);
            }
        }
    }

    internal sealed class CsvBoundColumnMap<T>
        where T : allows ref struct
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

    internal struct CsvRowProjector<T>
        where T : allows ref struct
    {
        private readonly TypeMapInfo<T> _typeInfo;
        private readonly StringComparer _comparer;
        private readonly HeaderNormalization _normalization;
        private readonly int _headerRow;
        private readonly IFormatProvider _provider;
        private readonly bool _throwOnParseFailure;
        private ColumnParser<T>[]? _fieldParsers;
        private string?[] _fieldNames;
        private bool[] _fieldRequired;
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

        private static bool IsBlank(CsvReader.Enumerator rows)
        {
            return rows.FieldCount == 1 && rows.FieldAt(0).Type == CellType.Empty;
        }

        internal ProjectionStep Classify(CsvReader.Enumerator rows)
        {
            if (_headerRow < 0)
            {
                _rowNumber++;
                return IsBlank(rows) ? ProjectionStep.Skip : ProjectionStep.Yield;
            }
            if (_typeInfo.IsIndexBased)
            {
                if (_fieldParsers is null)
                {
                    BuildIndexColumnMap();
                }
                _rowNumber++;
                return IsBlank(rows) ? ProjectionStep.Skip : ProjectionStep.Yield;
            }
            if (_fieldParsers is not null && _rowNumber >= _headerRow)
            {
                _rowNumber++;
                return IsBlank(rows) ? ProjectionStep.Skip : ProjectionStep.Yield;
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
            return IsBlank(rows) ? ProjectionStep.Skip : ProjectionStep.Yield;
        }

        internal readonly T Project(CsvReader.Enumerator rows)
        {
            T model = _typeInfo.CreateInstance();
            ParseCurrentRow(rows, ref model);
            return model;
        }

        private void BuildColumnMap(CsvReader.Enumerator rows)
        {
            CsvBoundColumnMap<T> map = BuildBoundMap(rows, _typeInfo, _comparer, _normalization);
            _fieldParsers = map.FieldParsers;
            _fieldNames = map.FieldNames;
            _fieldRequired = map.FieldRequired;
            _requiredFields = map.RequiredFields;
        }

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
