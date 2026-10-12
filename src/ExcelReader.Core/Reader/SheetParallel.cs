using ExcelReader.Core.Parser;
using ExcelReader.Core.Parser.Internal;
using ExcelReader.Core.Parser.ParallelCsv;
using ExcelReader.Core.Parser.SheetPipeline;
using ExcelReader.Core.Reader.Csv;

namespace ExcelReader.Core.Reader
{
    /// <summary>
    /// Reads a sheet of any format across several threads: <c>ParseAsync</c> yields typed rows in source
    /// order, <c>ForEachAsync</c> hands each row to a callback on the worker threads, and
    /// <c>AggregateAsync</c> folds rows into one accumulator per partition.
    /// </summary>
    /// <remarks>
    /// <para>
    /// The caller owns the workbook and keeps it open until the returned task completes or the
    /// enumeration ends. The CSV dialect is the <see cref="CsvReaderOptions"/> the workbook was opened with.
    /// </para>
    /// <para>
    /// A read falls back to one sequential pass, with the same results and no error, when
    /// <see cref="SheetParallelOptions.DegreeOfParallelism"/> is 1, the sheet is too small to partition
    /// usefully, the sheet is an XLS sheet or a table view, or a CSV cannot be partitioned (a stream that is not a file
    /// or a non-UTF-8 encoding).
    /// </para>
    /// <para>
    /// Delivery guarantees. The general contract is the weaker CSV one: <c>ForEachAsync</c> is at-least-once,
    /// and a partition whose start was guessed wrongly is discarded and read again. XLSB delivers every row
    /// exactly once, in <c>ForEachAsync</c> too. XLSX delivers exactly once for a well-formed sheet; a
    /// malformed one can make <c>ForEachAsync</c> deliver a row twice. <c>ParseAsync</c> yields the same items in
    /// the same order as the sequential <see cref="ExcelParser{T}"/> on every format, and
    /// <c>AggregateAsync</c> and <c>ForEachAsync</c> equal their own single-thread run.
    /// </para>
    /// <para>
    /// <see cref="ExcelParseException.Row"/>: <c>ParseAsync</c> renumbers it from the count emitted at the
    /// merge; <c>ForEachAsync</c> and <c>AggregateAsync</c> report the exact row when sequential and 0 when partitioned.
    /// </para>
    /// </remarks>
    public static class SheetParallel
    {
        /// <summary>
        /// Parses a sheet into <typeparamref name="T"/> instances across several threads, yielding the
        /// same sequence, in the same order, that the sequential <see cref="ExcelParser{T}"/> produces.
        /// </summary>
        /// <typeparam name="T">The row model type to bind each row to.</typeparam>
        /// <param name="sheet">The sheet to read. The caller keeps ownership and keeps its workbook open until the enumeration ends.</param>
        /// <param name="parser">How columns bind to <typeparamref name="T"/>; created with <see cref="ExcelParser"/>. Its <see cref="ExcelParserOptions.HeaderRow"/> locates the header.</param>
        /// <param name="options">Parallelism options. Defaults to <see cref="SheetParallelOptions.Default"/>. <see cref="SheetParallelOptions.HeaderRow"/> does not apply here and must be left at <c>0</c>.</param>
        /// <param name="ct">A token to cancel enumeration.</param>
        /// <returns>An enumerable that yields one <typeparamref name="T"/> per data row.</returns>
        /// <exception cref="ArgumentNullException"><paramref name="sheet"/> or <paramref name="parser"/> is <see langword="null"/>.</exception>
        /// <exception cref="ArgumentException"><see cref="SheetParallelOptions.HeaderRow"/> is not 0.</exception>
        /// <exception cref="ArgumentOutOfRangeException">A numeric option is negative.</exception>
        /// <remarks>
        /// Parsing and type conversion run in parallel; whatever the caller does per row does not. A caller
        /// whose per-row work dominates will see little gain from raising <see cref="SheetParallelOptions.DegreeOfParallelism"/>.
        /// For CSV, <see cref="CsvReaderOptions.InternStrings"/> yields substantially less here than sequentially,
        /// since each partition keeps its own dedup cache. A partition read from a wrongly guessed start is
        /// discarded before any of its rows are yielded, but its cells have already passed through the parser's
        /// converters and property setters, which must therefore be free of side effects.
        /// </remarks>
        public static IAsyncEnumerable<T> ParseAsync<T>(
            IExcelSheet sheet, ExcelParser<T> parser, SheetParallelOptions? options = null, CancellationToken ct = default)
        {
            ArgumentNullException.ThrowIfNull(sheet);
            ArgumentNullException.ThrowIfNull(parser);
            SheetParallelOptions validated = Validated(options);
            if (validated.HeaderRow != 0)
            {
                throw new ArgumentException(
                    $"{nameof(SheetParallelOptions)}.{nameof(SheetParallelOptions.HeaderRow)} does not apply to typed parsing; set {nameof(ExcelParserOptions)}.{nameof(ExcelParserOptions.HeaderRow)} instead.",
                    nameof(options));
            }
            return SheetRunner.Parse(sheet, parser, validated.DegreeOfParallelism, ct);
        }

        /// <summary>
        /// Reads a sheet across several threads, parses every row into a <typeparamref name="TRecord"/>
        /// and hands it to <paramref name="body"/>.
        /// </summary>
        /// <typeparam name="TRecord">The record type. A record whose <see cref="IRowRecord{TSelf}.TryParse"/> returns <see langword="false"/> is skipped.</typeparam>
        /// <param name="sheet">The sheet to read. The caller keeps ownership and keeps its workbook open until the task completes.</param>
        /// <param name="body">Receives each record. See the remarks for the contract it must honor.</param>
        /// <param name="options">Parallelism and header options. Defaults to <see cref="SheetParallelOptions.Default"/>.</param>
        /// <param name="ct">A token to cancel processing.</param>
        /// <returns>A task that completes when every record has been delivered.</returns>
        /// <exception cref="ArgumentNullException"><paramref name="sheet"/> or <paramref name="body"/> is <see langword="null"/>.</exception>
        /// <exception cref="ArgumentOutOfRangeException">A numeric option is negative.</exception>
        /// <remarks>
        /// <para>
        /// <paramref name="body"/> runs concurrently on worker threads and must be safe to call from
        /// several at once; the caller owns any synchronization. Records arrive in no particular order.
        /// </para>
        /// <para>
        /// Delivery is at least once in general: <paramref name="body"/> may be invoked more than once for the
        /// same record. On CSV, a partition whose start was guessed wrongly is read again from the confirmed
        /// offset, and the discarded speculative pass may already have delivered a whole partition's worth of
        /// records, which may also be misparsed: a wrong start shifts every field boundary. An exception
        /// <paramref name="body"/> throws during a discarded pass is discarded with it. A start is only guessed
        /// wrongly inside a quoted field, so a CSV in which the quote character never appears delivers every
        /// record exactly once. XLSB delivers every row exactly once. XLSX does too for a well-formed sheet; on
        /// a malformed one, workers may already have called <paramref name="body"/> for rows that are then
        /// dropped, so a row can be delivered twice or be one the sequential reader would not produce.
        /// </para>
        /// <para>
        /// A <paramref name="body"/> with side effects must therefore be idempotent and tolerate records that
        /// do not exist. For exactly-once on any source, use
        /// <see cref="AggregateAsync{TAccumulator, TRecord}(IExcelSheet, SheetParallelOptions?, CancellationToken)"/>, where the
        /// discarded partition's accumulator is thrown away, and apply the side effects once it returns.
        /// </para>
        /// <para>
        /// A <see cref="ReadOnlySpan{T}"/> column points into a pooled buffer the worker reuses once
        /// <paramref name="body"/> returns. Copy anything that must outlive the call.
        /// </para>
        /// </remarks>
        public static Task ForEachAsync<TRecord>(
            IExcelSheet sheet, Action<TRecord> body, SheetParallelOptions? options = null, CancellationToken ct = default)
            where TRecord : IRowRecord<TRecord>, allows ref struct
        {
            ArgumentNullException.ThrowIfNull(sheet);
            ArgumentNullException.ThrowIfNull(body);
            return SheetRunner.RunAsync(sheet, RecordCallback<TRecord>.For(body), null, Validated(options), ct);
        }

        /// <summary>
        /// Reads a sheet across several threads, binds every row to a <typeparamref name="TModel"/>
        /// through <paramref name="parser"/> and hands it to <paramref name="body"/>.
        /// </summary>
        /// <typeparam name="TModel">The model type. May be a <see langword="ref struct"/> with <see cref="ReadOnlySpan{T}"/> of <see cref="byte"/> properties.</typeparam>
        /// <param name="sheet">The sheet to read. The caller keeps ownership and keeps its workbook open until the task completes.</param>
        /// <param name="parser">How columns bind to <typeparamref name="TModel"/>; created with <see cref="ExcelParser"/>. Its <see cref="ExcelParserOptions.HeaderRow"/> is ignored in favour of <see cref="SheetParallelOptions.HeaderRow"/>.</param>
        /// <param name="body">Receives each record. See the remarks for the contract it must honor.</param>
        /// <param name="options">Parallelism and header options. Defaults to <see cref="SheetParallelOptions.Default"/>.</param>
        /// <param name="ct">A token to cancel processing.</param>
        /// <returns>A task that completes when every record has been delivered.</returns>
        /// <exception cref="ArgumentNullException"><paramref name="sheet"/>, <paramref name="parser"/> or <paramref name="body"/> is <see langword="null"/>.</exception>
        /// <exception cref="ArgumentException"><paramref name="parser"/> binds columns by header name and <see cref="SheetParallelOptions.HeaderRow"/> is 0.</exception>
        /// <exception cref="ArgumentOutOfRangeException">A numeric option is negative.</exception>
        /// <remarks>
        /// <para>
        /// Carries the same concurrency, delivery and span-lifetime contract as the
        /// <see cref="IRowRecord{TSelf}"/> overload: <paramref name="body"/> runs concurrently, delivery is
        /// at least once in general (exactly once on XLSB, and on a well-formed XLSX sheet), and records arrive
        /// in no particular order. A property that fails to parse keeps its default unless
        /// <see cref="ExcelParserOptions.ThrowOnParseFailure"/> is set; a missing <c>[ExcelRequired]</c> column or
        /// value throws <see cref="ExcelParseException"/>; empty rows are skipped.
        /// </para>
        /// <para>
        /// <see cref="ExcelParseException.Row"/> is exact when the sheet is read sequentially and
        /// 0 when it is partitioned, since a partition does not know how many rows precede it.
        /// </para>
        /// </remarks>
        public static Task ForEachAsync<TModel>(
            IExcelSheet sheet, ExcelParser<TModel> parser, Action<TModel> body, SheetParallelOptions? options = null, CancellationToken ct = default)
            where TModel : allows ref struct
        {
            ArgumentNullException.ThrowIfNull(sheet);
            ArgumentNullException.ThrowIfNull(parser);
            ArgumentNullException.ThrowIfNull(body);
            SheetParallelOptions validated = Validated(options);
            TypeMapInfo<TModel> info = InfoFor(sheet, parser, validated);
            return SheetRunner.RunAsync(
                sheet, MappedCallback<TModel>.Unbound, MappedCallback<TModel>.Binder(parser, info, sheet.IsDate1904, validated.HeaderRow, body), validated, ct);
        }

        /// <summary>
        /// Reads a sheet across several threads, parses every row into a <typeparamref name="TRecord"/>
        /// and folds it into a <typeparamref name="TAccumulator"/>, one instance per partition, merged in sheet order.
        /// </summary>
        /// <typeparam name="TAccumulator">The accumulator type. See <see cref="IRowAccumulator{TSelf, TModel}"/> for the contract it must honor.</typeparam>
        /// <typeparam name="TRecord">The record type. A record whose <see cref="IRowRecord{TSelf}.TryParse"/> returns <see langword="false"/> is skipped.</typeparam>
        /// <param name="sheet">The sheet to read. The caller keeps ownership and keeps its workbook open until the task completes.</param>
        /// <param name="options">Parallelism and header options. Defaults to <see cref="SheetParallelOptions.Default"/>.</param>
        /// <param name="ct">A token to cancel processing.</param>
        /// <returns>The accumulator holding every record.</returns>
        /// <exception cref="ArgumentNullException"><paramref name="sheet"/> is <see langword="null"/>.</exception>
        /// <exception cref="ArgumentOutOfRangeException">A numeric option is negative.</exception>
        /// <remarks>
        /// A partition read from a wrongly guessed start is read again into a new accumulator. The result
        /// holds every record exactly once, but <see cref="IRowRecord{TSelf}.TryParse"/> and
        /// <see cref="IRowAccumulator{TSelf, TModel}.Add"/> may already have run for the discarded pass's records, as
        /// <see cref="IRowAccumulator{TSelf, TModel}"/> describes. A sequential fallback uses a single accumulator and never calls
        /// <see cref="IRowAccumulator{TSelf, TModel}.Merge"/>.
        /// </remarks>
        public static Task<TAccumulator> AggregateAsync<TAccumulator, TRecord>(
            IExcelSheet sheet, SheetParallelOptions? options = null, CancellationToken ct = default)
            where TAccumulator : IRowAccumulator<TAccumulator, TRecord>, new()
            where TRecord : IRowRecord<TRecord>, allows ref struct
        {
            ArgumentNullException.ThrowIfNull(sheet);
            return SheetRunner.RunAsync(sheet, RecordAggregation<TAccumulator, TRecord>.Instance, null, Validated(options), ct);
        }

        /// <summary>
        /// Reads a sheet across several threads, binds every row to a <typeparamref name="TModel"/>
        /// through <paramref name="parser"/> and folds it into a <typeparamref name="TAccumulator"/>, one
        /// instance per partition, merged in sheet order.
        /// </summary>
        /// <typeparam name="TAccumulator">The accumulator type. See <see cref="IRowAccumulator{TSelf, TModel}"/> for the contract it must honor.</typeparam>
        /// <typeparam name="TModel">The model type.</typeparam>
        /// <param name="sheet">The sheet to read. The caller keeps ownership and keeps its workbook open until the task completes.</param>
        /// <param name="parser">How columns bind to <typeparamref name="TModel"/>; created with <see cref="ExcelParser"/>. Its <see cref="ExcelParserOptions.HeaderRow"/> is ignored in favour of <see cref="SheetParallelOptions.HeaderRow"/>.</param>
        /// <param name="options">Parallelism and header options. Defaults to <see cref="SheetParallelOptions.Default"/>.</param>
        /// <param name="ct">A token to cancel processing.</param>
        /// <returns>The accumulator holding every record.</returns>
        /// <exception cref="ArgumentNullException"><paramref name="sheet"/> or <paramref name="parser"/> is <see langword="null"/>.</exception>
        /// <exception cref="ArgumentException"><paramref name="parser"/> binds columns by header name and <see cref="SheetParallelOptions.HeaderRow"/> is 0.</exception>
        /// <exception cref="ArgumentOutOfRangeException">A numeric option is negative.</exception>
        /// <remarks>
        /// <para>
        /// The header is bound once, before any row is folded. A property that fails to parse keeps its
        /// default unless <see cref="ExcelParserOptions.ThrowOnParseFailure"/> is set; a missing
        /// <c>[ExcelRequired]</c> column or value throws <see cref="ExcelParseException"/>; empty rows are skipped.
        /// </para>
        /// <para>
        /// <see cref="ExcelParseException.Row"/> is exact when the sheet is read sequentially and 0 when it is
        /// partitioned, since a partition does not know how many rows precede it.
        /// </para>
        /// <para>Otherwise carries the same contract and fallbacks as the <see cref="IRowRecord{TSelf}"/> overload.</para>
        /// </remarks>
        public static Task<TAccumulator> AggregateAsync<TAccumulator, TModel>(
            IExcelSheet sheet, ExcelParser<TModel> parser, SheetParallelOptions? options = null, CancellationToken ct = default)
            where TAccumulator : IRowAccumulator<TAccumulator, TModel>, new()
            where TModel : allows ref struct
        {
            ArgumentNullException.ThrowIfNull(sheet);
            ArgumentNullException.ThrowIfNull(parser);
            SheetParallelOptions validated = Validated(options);
            TypeMapInfo<TModel> info = InfoFor(sheet, parser, validated);
            return SheetRunner.RunAsync(
                sheet, MappedAggregation<TAccumulator, TModel>.Unbound,
                MappedAggregation<TAccumulator, TModel>.Binder(parser, info, sheet.IsDate1904, validated.HeaderRow), validated, ct);
        }

        /// <summary>
        /// Reads a sheet across several threads and folds every row with the functions of
        /// <paramref name="aggregation"/>, one accumulator per partition, combined in sheet order.
        /// </summary>
        /// <typeparam name="TState">The accumulator type.</typeparam>
        /// <param name="sheet">The sheet to read. The caller keeps ownership and keeps its workbook open until the task completes.</param>
        /// <param name="aggregation">The seed, accumulate and combine functions. They carry the same contract as <see cref="IRowAccumulator{TSelf, TModel}"/>.</param>
        /// <param name="options">Parallelism and header options. Defaults to <see cref="SheetParallelOptions.Default"/>.</param>
        /// <param name="ct">A token to cancel processing.</param>
        /// <returns>The combined accumulator.</returns>
        /// <exception cref="ArgumentNullException"><paramref name="sheet"/> or <paramref name="aggregation"/>, or one of its functions, is <see langword="null"/>.</exception>
        /// <exception cref="ArgumentOutOfRangeException">A numeric option is negative.</exception>
        /// <remarks>Carries the same fallbacks as the <see cref="IRowAccumulator{TSelf, TModel}"/> overloads.</remarks>
        public static Task<TState> AggregateAsync<TState>(
            IExcelSheet sheet, RowAggregation<TState> aggregation, SheetParallelOptions? options = null, CancellationToken ct = default)
        {
            ArgumentNullException.ThrowIfNull(sheet);
            ArgumentNullException.ThrowIfNull(aggregation);
            ArgumentNullException.ThrowIfNull(aggregation.Seed, nameof(aggregation));
            ArgumentNullException.ThrowIfNull(aggregation.Accumulate, nameof(aggregation));
            ArgumentNullException.ThrowIfNull(aggregation.Combine, nameof(aggregation));
            return SheetRunner.RunAsync(sheet, aggregation, null, Validated(options), ct);
        }

        private static TypeMapInfo<TModel> InfoFor<TModel>(IExcelSheet sheet, ExcelParser<TModel> parser, SheetParallelOptions options)
            where TModel : allows ref struct
        {
            TypeMapInfo<TModel> info = sheet is CsvSheet ? parser.CsvInfo : parser.Info;
            if (options.HeaderRow == 0 && !info.IsIndexBased)
            {
                throw new ArgumentException(
                    $"A parser that binds columns by header name needs {nameof(SheetParallelOptions)}.{nameof(SheetParallelOptions.HeaderRow)} of at least 1.", nameof(parser));
            }
            return info;
        }

        private static SheetParallelOptions Validated(SheetParallelOptions? options)
        {
            options ??= SheetParallelOptions.Default;
            ArgumentOutOfRangeException.ThrowIfNegative(options.DegreeOfParallelism, nameof(options));
            ArgumentOutOfRangeException.ThrowIfNegative(options.HeaderRow, nameof(options));
            return options;
        }
    }
}
