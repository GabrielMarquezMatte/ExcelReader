using ExcelReader.Core.Parser;
using ExcelReader.Core.Parser.ParallelCsv;

namespace ExcelReader.Core.Reader.Csv
{
    /// <summary>
    /// Reads CSV across several threads: <see cref="ParseAsync{T}(string, ExcelParser{T}, CsvParallelOptions?, CancellationToken)"/>
    /// yields typed rows in source order, <c>ForEachAsync</c> hands each record to a callback on the worker
    /// threads, and <c>AggregateAsync</c> folds records into one accumulator per partition.
    /// </summary>
    public static class CsvParallel
    {
        /// <summary>
        /// Parses a CSV file into <typeparamref name="T"/> instances across several threads, yielding the
        /// same sequence, in the same order, that the sequential <see cref="ExcelParser{T}"/> produces.
        /// </summary>
        /// <typeparam name="T">The row model type to bind each CSV record to.</typeparam>
        /// <param name="path">The path of the CSV file to read.</param>
        /// <param name="parser">How columns bind to <typeparamref name="T"/>; created with <see cref="ExcelParser"/>. Its <see cref="ExcelParserConfig.HeaderRow"/> locates the header.</param>
        /// <param name="options">Parallelism and dialect options. Defaults to <see cref="CsvParallelOptions.Default"/>. <see cref="CsvParallelOptions.HeaderRow"/> does not apply here and must be left at <c>0</c>.</param>
        /// <param name="ct">A token to cancel enumeration.</param>
        /// <remarks>
        /// <para>
        /// Parsing and type conversion run in parallel; whatever the caller does per row does not. A caller
        /// whose per-row work dominates will see little gain from raising <see cref="CsvParallelOptions.DegreeOfParallelism"/>.
        /// </para>
        /// <para>
        /// Falls back to sequential parsing — same results, one thread — when the source is too small to
        /// partition usefully, or when <see cref="CsvReaderOptions.Encoding"/> is set to a non-UTF-8 encoding,
        /// which must be transcoded sequentially.
        /// </para>
        /// <para>
        /// <see cref="CsvReaderOptions.InternStrings"/> yields substantially less here than on the sequential
        /// path: each partition keeps its own dedup cache, so hit rates fall and memory multiplies. Results
        /// are unaffected.
        /// </para>
        /// <para>
        /// Each row is yielded exactly once. A partition read from a wrongly guessed start is discarded before
        /// any of its rows are yielded, but its cells have already passed through the parser's converters and
        /// property setters, which must therefore be free of side effects.
        /// </para>
        /// </remarks>
        public static IAsyncEnumerable<T> ParseAsync<T>(
            string path,
            ExcelParser<T> parser,
            CsvParallelOptions? options = null,
            CancellationToken ct = default)
        {
            ArgumentNullException.ThrowIfNull(parser);
            ArgumentException.ThrowIfNullOrEmpty(path);
            CsvParallelOptions parallel = ValidatedForTypedParse(options);
            return ParallelCsvFactory.Create(path, parallel.DegreeOfParallelism, parallel.Reader, parser, ct);
        }

        /// <summary>
        /// Parses an in-memory CSV buffer into <typeparamref name="T"/> instances across several threads,
        /// yielding the same sequence, in the same order, that the sequential <see cref="ExcelParser{T}"/> produces.
        /// </summary>
        /// <typeparam name="T">The row model type to bind each CSV record to.</typeparam>
        /// <param name="data">The CSV bytes. The caller keeps ownership; the buffer must not be mutated during enumeration.</param>
        /// <param name="parser">How columns bind to <typeparamref name="T"/>; created with <see cref="ExcelParser"/>. Its <see cref="ExcelParserConfig.HeaderRow"/> locates the header.</param>
        /// <param name="options">Parallelism and dialect options. Defaults to <see cref="CsvParallelOptions.Default"/>. <see cref="CsvParallelOptions.HeaderRow"/> does not apply here and must be left at <c>0</c>.</param>
        /// <param name="ct">A token to cancel enumeration.</param>
        /// <remarks>Carries the same fallback, <see cref="CsvReaderOptions.InternStrings"/> and exactly-once caveats as the path-based overload.</remarks>
        public static IAsyncEnumerable<T> ParseAsync<T>(
            ReadOnlyMemory<byte> data,
            ExcelParser<T> parser,
            CsvParallelOptions? options = null,
            CancellationToken ct = default)
        {
            ArgumentNullException.ThrowIfNull(parser);
            CsvParallelOptions parallel = ValidatedForTypedParse(options);
            return ParallelCsvFactory.Create(data, parallel.DegreeOfParallelism, parallel.Reader, parser, ct);
        }

        /// <summary>
        /// Parses a CSV stream into <typeparamref name="T"/> instances, in parallel where the stream can be
        /// partitioned, yielding the same sequence, in the same order, that the sequential
        /// <see cref="ExcelParser{T}"/> produces.
        /// </summary>
        /// <typeparam name="T">The row model type to bind each CSV record to.</typeparam>
        /// <param name="stream">The CSV stream, read from its current position. The caller keeps ownership and must not read from it concurrently.</param>
        /// <param name="parser">How columns bind to <typeparamref name="T"/>; created with <see cref="ExcelParser"/>. Its <see cref="ExcelParserConfig.HeaderRow"/> locates the header.</param>
        /// <param name="options">Parallelism and dialect options. Defaults to <see cref="CsvParallelOptions.Default"/>. <see cref="CsvParallelOptions.HeaderRow"/> does not apply here and must be left at <c>0</c>.</param>
        /// <param name="ct">A token to cancel enumeration.</param>
        /// <remarks>
        /// <para>
        /// Only a <see cref="FileStream"/>, or a <see cref="MemoryStream"/> whose buffer is publicly
        /// visible, can be partitioned; every other stream — including seekable ones — parses sequentially,
        /// with identical results. A stream exposes a single mutable position, and a seek can be
        /// arbitrarily expensive, so this overload partitions only what it can read positionally and
        /// cheaply rather than promising parallelism it cannot deliver.
        /// </para>
        /// <para>
        /// Passing a <see cref="FileStream"/> reads its <see cref="FileStream.SafeFileHandle"/>, which
        /// flushes that stream's internal buffer and disables its subsequent buffering optimizations. The
        /// stream's position is not moved. Prefer the path-based overload where a path is available.
        /// </para>
        /// <para>Carries the same <see cref="CsvReaderOptions.InternStrings"/> and exactly-once caveats as the other overloads.</para>
        /// </remarks>
        public static IAsyncEnumerable<T> ParseAsync<T>(
            Stream stream,
            ExcelParser<T> parser,
            CsvParallelOptions? options = null,
            CancellationToken ct = default)
        {
            ArgumentNullException.ThrowIfNull(parser);
            ArgumentNullException.ThrowIfNull(stream);
            CsvParallelOptions parallel = ValidatedForTypedParse(options);
            return ParallelCsvFactory.Create(stream, parallel.DegreeOfParallelism, parallel.Reader, parser, ct);
        }

        /// <summary>
        /// Reads an in-memory CSV buffer across several threads, parses every record into a
        /// <typeparamref name="TRecord"/> and folds it into a <typeparamref name="TAccumulator"/>, one
        /// instance per partition, merged in buffer order.
        /// </summary>
        /// <typeparam name="TAccumulator">The accumulator type. See <see cref="ICsvAccumulator{TSelf, TModel}"/> for the contract it must honor.</typeparam>
        /// <typeparam name="TRecord">The record type. A record whose <see cref="ICsvRecord{TSelf}.TryParse"/> returns <see langword="false"/> is skipped.</typeparam>
        /// <param name="data">The CSV bytes. The caller keeps ownership; the buffer must not be mutated during processing.</param>
        /// <param name="options">Parallelism, dialect and header options. Defaults to <see cref="CsvParallelOptions.Default"/>.</param>
        /// <param name="ct">A token to cancel processing.</param>
        /// <returns>The accumulator holding every record.</returns>
        /// <remarks>
        /// <para>
        /// Partitions start at guessed record boundaries; one whose guess landed inside a quoted field
        /// spanning lines is read again into a new accumulator. The result holds every record exactly once,
        /// but <see cref="ICsvRecord{TSelf}.TryParse"/> and <see cref="ICsvAccumulator{TSelf, TModel}.Add"/>
        /// may already have run for the discarded pass's records, as <see cref="ICsvAccumulator{TSelf, TModel}"/> describes.
        /// </para>
        /// <para>
        /// Falls back to one sequential pass, with a single accumulator and no call to
        /// <see cref="ICsvAccumulator{TSelf, TModel}.Merge"/>, when the source is too small to partition usefully
        /// or when <see cref="CsvReaderOptions.Encoding"/> is set to a non-UTF-8 encoding.
        /// </para>
        /// </remarks>
        public static Task<TAccumulator> AggregateAsync<TAccumulator, TRecord>(
            ReadOnlyMemory<byte> data,
            CsvParallelOptions? options = null,
            CancellationToken ct = default)
            where TAccumulator : ICsvAccumulator<TAccumulator, TRecord>, new()
            where TRecord : ICsvRecord<TRecord>, allows ref struct
        {
            return ParallelCsvProcessor.RunAsync(data, RecordAggregation<TAccumulator, TRecord>.Instance, null, Validated(options), ct);
        }

        /// <summary>
        /// Reads a CSV file across several threads, parses every record into a <typeparamref name="TRecord"/>
        /// and folds it into a <typeparamref name="TAccumulator"/>, one instance per partition, merged in file order.
        /// </summary>
        /// <typeparam name="TAccumulator">The accumulator type. See <see cref="ICsvAccumulator{TSelf, TModel}"/> for the contract it must honor.</typeparam>
        /// <typeparam name="TRecord">The record type. A record whose <see cref="ICsvRecord{TSelf}.TryParse"/> returns <see langword="false"/> is skipped.</typeparam>
        /// <param name="path">The path of the CSV file to read.</param>
        /// <param name="options">Parallelism, dialect and header options. Defaults to <see cref="CsvParallelOptions.Default"/>.</param>
        /// <param name="ct">A token to cancel processing.</param>
        /// <returns>The accumulator holding every record.</returns>
        /// <remarks>Carries the same contract and fallbacks as the in-memory overload.</remarks>
        public static Task<TAccumulator> AggregateAsync<TAccumulator, TRecord>(
            string path,
            CsvParallelOptions? options = null,
            CancellationToken ct = default)
            where TAccumulator : ICsvAccumulator<TAccumulator, TRecord>, new()
            where TRecord : ICsvRecord<TRecord>, allows ref struct
        {
            ArgumentException.ThrowIfNullOrEmpty(path);
            return ParallelCsvProcessor.RunAsync(path, RecordAggregation<TAccumulator, TRecord>.Instance, null, Validated(options), ct);
        }

        /// <summary>
        /// Reads a CSV stream, in parallel where the stream can be partitioned, parses every record into a
        /// <typeparamref name="TRecord"/> and folds it into a <typeparamref name="TAccumulator"/>, one instance
        /// per partition, merged in stream order.
        /// </summary>
        /// <typeparam name="TAccumulator">The accumulator type. See <see cref="ICsvAccumulator{TSelf, TModel}"/> for the contract it must honor.</typeparam>
        /// <typeparam name="TRecord">The record type. A record whose <see cref="ICsvRecord{TSelf}.TryParse"/> returns <see langword="false"/> is skipped.</typeparam>
        /// <param name="stream">The CSV stream, read from its current position. The caller keeps ownership and must not read from it concurrently.</param>
        /// <param name="options">Parallelism, dialect and header options. Defaults to <see cref="CsvParallelOptions.Default"/>.</param>
        /// <param name="ct">A token to cancel processing.</param>
        /// <returns>The accumulator holding every record.</returns>
        /// <remarks>
        /// Carries the same contract and fallbacks as the in-memory overload, and partitions only a
        /// <see cref="FileStream"/> or a <see cref="MemoryStream"/> whose buffer is publicly visible;
        /// every other stream is read sequentially.
        /// </remarks>
        public static Task<TAccumulator> AggregateAsync<TAccumulator, TRecord>(
            Stream stream,
            CsvParallelOptions? options = null,
            CancellationToken ct = default)
            where TAccumulator : ICsvAccumulator<TAccumulator, TRecord>, new()
            where TRecord : ICsvRecord<TRecord>, allows ref struct
        {
            ArgumentNullException.ThrowIfNull(stream);
            return ParallelCsvProcessor.RunAsync(stream, RecordAggregation<TAccumulator, TRecord>.Instance, null, Validated(options), ct);
        }

        /// <summary>
        /// Reads a CSV file across several threads, parses every record into a <typeparamref name="TRecord"/>
        /// and hands it to <paramref name="body"/>.
        /// </summary>
        /// <typeparam name="TRecord">The record type. A record whose <see cref="ICsvRecord{TSelf}.TryParse"/> returns <see langword="false"/> is skipped.</typeparam>
        /// <param name="path">The path of the CSV file to read.</param>
        /// <param name="body">Receives each record. See the remarks for the contract it must honor.</param>
        /// <param name="options">Parallelism, dialect and header options. Defaults to <see cref="CsvParallelOptions.Default"/>.</param>
        /// <param name="ct">A token to cancel processing.</param>
        /// <returns>A task that completes when every record has been delivered.</returns>
        /// <exception cref="ArgumentNullException"><paramref name="body"/> is <see langword="null"/>.</exception>
        /// <remarks>
        /// <para>
        /// <paramref name="body"/> runs concurrently on worker threads and must be safe to call from
        /// several at once; the caller owns any synchronization. Records arrive in no particular order.
        /// </para>
        /// <para>
        /// Delivery is at least once, not exactly once: <paramref name="body"/> may be invoked more than once
        /// for the same record. A partition whose start was guessed wrongly is read again from the confirmed
        /// offset, and the discarded speculative pass may already have delivered a whole partition's worth of
        /// records — 1 MiB to 64 MiB of them, not just a few near the seam. Those records may also be misparsed:
        /// a wrong start shifts every field boundary, so a quoted field holding a delimiter can split into values
        /// that appear nowhere in the source. An exception <paramref name="body"/> throws during a discarded
        /// pass is discarded with it.
        /// </para>
        /// <para>
        /// A <paramref name="body"/> with side effects must therefore be idempotent, and must also tolerate
        /// records that do not exist; deduplicating on a key does not filter those out. A start is only guessed
        /// wrongly when it falls inside a quoted field, so a source in which the quote character never appears
        /// delivers every record exactly once. For that guarantee on any source, use
        /// <see cref="AggregateAsync{TAccumulator, TRecord}(string, CsvParallelOptions?, CancellationToken)"/>,
        /// where the discarded partition's accumulator is thrown away, and apply the side effects once it returns.
        /// </para>
        /// <para>
        /// A <see cref="ReadOnlySpan{T}"/> column points into a pooled buffer the worker reuses once
        /// <paramref name="body"/> returns. Copy anything that must outlive the call.
        /// </para>
        /// <para>
        /// Falls back to one sequential pass, delivering records once each in source order on a single
        /// thread, when the source is too small to partition usefully or when
        /// <see cref="CsvReaderOptions.Encoding"/> is set to a non-UTF-8 encoding.
        /// </para>
        /// </remarks>
        public static Task ForEachAsync<TRecord>(
            string path,
            Action<TRecord> body,
            CsvParallelOptions? options = null,
            CancellationToken ct = default)
            where TRecord : ICsvRecord<TRecord>, allows ref struct
        {
            ArgumentException.ThrowIfNullOrEmpty(path);
            ArgumentNullException.ThrowIfNull(body);
            return ParallelCsvProcessor.RunAsync(path, RecordCallback<TRecord>.For(body), null, Validated(options), ct);
        }

        /// <summary>
        /// Reads an in-memory CSV buffer across several threads, parses every record into a
        /// <typeparamref name="TRecord"/> and hands it to <paramref name="body"/>.
        /// </summary>
        /// <typeparam name="TRecord">The record type. A record whose <see cref="ICsvRecord{TSelf}.TryParse"/> returns <see langword="false"/> is skipped.</typeparam>
        /// <param name="data">The CSV bytes. The caller keeps ownership; the buffer must not be mutated during processing.</param>
        /// <param name="body">Receives each record. See the remarks for the contract it must honor.</param>
        /// <param name="options">Parallelism, dialect and header options. Defaults to <see cref="CsvParallelOptions.Default"/>.</param>
        /// <param name="ct">A token to cancel processing.</param>
        /// <returns>A task that completes when every record has been delivered.</returns>
        /// <exception cref="ArgumentNullException"><paramref name="body"/> is <see langword="null"/>.</exception>
        /// <remarks>Carries the same contract and fallbacks as the path-based overload.</remarks>
        public static Task ForEachAsync<TRecord>(
            ReadOnlyMemory<byte> data,
            Action<TRecord> body,
            CsvParallelOptions? options = null,
            CancellationToken ct = default)
            where TRecord : ICsvRecord<TRecord>, allows ref struct
        {
            ArgumentNullException.ThrowIfNull(body);
            return ParallelCsvProcessor.RunAsync(data, RecordCallback<TRecord>.For(body), null, Validated(options), ct);
        }

        /// <summary>
        /// Reads a CSV stream, in parallel where the stream can be partitioned, parses every record into a
        /// <typeparamref name="TRecord"/> and hands it to <paramref name="body"/>.
        /// </summary>
        /// <typeparam name="TRecord">The record type. A record whose <see cref="ICsvRecord{TSelf}.TryParse"/> returns <see langword="false"/> is skipped.</typeparam>
        /// <param name="stream">The CSV stream, read from its current position. The caller keeps ownership and must not read from it concurrently.</param>
        /// <param name="body">Receives each record. See the remarks for the contract it must honor.</param>
        /// <param name="options">Parallelism, dialect and header options. Defaults to <see cref="CsvParallelOptions.Default"/>.</param>
        /// <param name="ct">A token to cancel processing.</param>
        /// <returns>A task that completes when every record has been delivered.</returns>
        /// <exception cref="ArgumentNullException"><paramref name="stream"/> or <paramref name="body"/> is <see langword="null"/>.</exception>
        /// <remarks>
        /// Carries the same contract and fallbacks as the path-based overload, and partitions only a
        /// <see cref="FileStream"/> or a <see cref="MemoryStream"/> whose buffer is publicly visible;
        /// every other stream is read sequentially.
        /// </remarks>
        public static Task ForEachAsync<TRecord>(
            Stream stream,
            Action<TRecord> body,
            CsvParallelOptions? options = null,
            CancellationToken ct = default)
            where TRecord : ICsvRecord<TRecord>, allows ref struct
        {
            ArgumentNullException.ThrowIfNull(stream);
            ArgumentNullException.ThrowIfNull(body);
            return ParallelCsvProcessor.RunAsync(stream, RecordCallback<TRecord>.For(body), null, Validated(options), ct);
        }

        /// <summary>
        /// Reads a CSV file across several threads, binds every record to a <typeparamref name="TModel"/>
        /// through <paramref name="parser"/> and hands it to <paramref name="body"/>.
        /// </summary>
        /// <typeparam name="TModel">The model type. May be a <see langword="ref struct"/> with <see cref="ReadOnlySpan{T}"/> of <see cref="byte"/> properties.</typeparam>
        /// <param name="path">The path of the CSV file to read.</param>
        /// <param name="parser">How columns bind to <typeparamref name="TModel"/>; created with <see cref="ExcelParser"/>. Its <see cref="ExcelParserConfig.HeaderRow"/> is ignored in favour of <see cref="CsvParallelOptions.HeaderRow"/>.</param>
        /// <param name="body">Receives each record. See the remarks for the contract it must honor.</param>
        /// <param name="options">Parallelism, dialect and header options. Defaults to <see cref="CsvParallelOptions.Default"/>.</param>
        /// <param name="ct">A token to cancel processing.</param>
        /// <returns>A task that completes when every record has been delivered.</returns>
        /// <exception cref="ArgumentNullException"><paramref name="parser"/> or <paramref name="body"/> is <see langword="null"/>.</exception>
        /// <exception cref="ArgumentException"><paramref name="parser"/> binds columns by header name and <see cref="CsvParallelOptions.HeaderRow"/> is 0.</exception>
        /// <remarks>
        /// <para>
        /// Carries the same concurrency, repeat-invocation and span-lifetime contract as the
        /// <see cref="ICsvRecord{TSelf}"/> overloads. A property that fails to parse keeps its default
        /// unless <see cref="ExcelParserConfig.ThrowOnParseFailure"/> is set; a missing
        /// <c>[ExcelRequired]</c> column or value throws <see cref="ExcelParseException"/>; empty
        /// records are skipped.
        /// </para>
        /// <para>
        /// <see cref="ExcelParseException.Row"/> is exact when the source is read sequentially and
        /// 0 when it is partitioned, since a partition does not know how many records precede it.
        /// </para>
        /// </remarks>
        public static Task ForEachAsync<TModel>(
            string path,
            ExcelParser<TModel> parser,
            Action<TModel> body,
            CsvParallelOptions? options = null,
            CancellationToken ct = default)
            where TModel : allows ref struct
        {
            ArgumentException.ThrowIfNullOrEmpty(path);
            CsvParallelOptions validated = Validated(options);
            return ParallelCsvProcessor.RunAsync(path, MappedCallback<TModel>.Unbound, BoundCallback(parser, body, validated), validated, ct);
        }

        /// <summary>
        /// Reads an in-memory CSV buffer across several threads, binds every record to a
        /// <typeparamref name="TModel"/> through <paramref name="parser"/> and hands it to <paramref name="body"/>.
        /// </summary>
        /// <typeparam name="TModel">The model type. May be a <see langword="ref struct"/> with <see cref="ReadOnlySpan{T}"/> of <see cref="byte"/> properties.</typeparam>
        /// <param name="data">The CSV bytes. The caller keeps ownership; the buffer must not be mutated during processing.</param>
        /// <param name="parser">How columns bind to <typeparamref name="TModel"/>; created with <see cref="ExcelParser"/>. Its <see cref="ExcelParserConfig.HeaderRow"/> is ignored in favour of <see cref="CsvParallelOptions.HeaderRow"/>.</param>
        /// <param name="body">Receives each record. See the remarks for the contract it must honor.</param>
        /// <param name="options">Parallelism, dialect and header options. Defaults to <see cref="CsvParallelOptions.Default"/>.</param>
        /// <param name="ct">A token to cancel processing.</param>
        /// <returns>A task that completes when every record has been delivered.</returns>
        /// <exception cref="ArgumentNullException"><paramref name="parser"/> or <paramref name="body"/> is <see langword="null"/>.</exception>
        /// <exception cref="ArgumentException"><paramref name="parser"/> binds columns by header name and <see cref="CsvParallelOptions.HeaderRow"/> is 0.</exception>
        /// <remarks>Carries the same contract and fallbacks as the path-based overload.</remarks>
        public static Task ForEachAsync<TModel>(
            ReadOnlyMemory<byte> data,
            ExcelParser<TModel> parser,
            Action<TModel> body,
            CsvParallelOptions? options = null,
            CancellationToken ct = default)
            where TModel : allows ref struct
        {
            CsvParallelOptions validated = Validated(options);
            return ParallelCsvProcessor.RunAsync(data, MappedCallback<TModel>.Unbound, BoundCallback(parser, body, validated), validated, ct);
        }

        /// <summary>
        /// Reads a CSV stream, in parallel where the stream can be partitioned, binds every record to a
        /// <typeparamref name="TModel"/> through <paramref name="parser"/> and hands it to <paramref name="body"/>.
        /// </summary>
        /// <typeparam name="TModel">The model type. May be a <see langword="ref struct"/> with <see cref="ReadOnlySpan{T}"/> of <see cref="byte"/> properties.</typeparam>
        /// <param name="stream">The CSV stream, read from its current position. The caller keeps ownership and must not read from it concurrently.</param>
        /// <param name="parser">How columns bind to <typeparamref name="TModel"/>; created with <see cref="ExcelParser"/>. Its <see cref="ExcelParserConfig.HeaderRow"/> is ignored in favour of <see cref="CsvParallelOptions.HeaderRow"/>.</param>
        /// <param name="body">Receives each record. See the remarks for the contract it must honor.</param>
        /// <param name="options">Parallelism, dialect and header options. Defaults to <see cref="CsvParallelOptions.Default"/>.</param>
        /// <param name="ct">A token to cancel processing.</param>
        /// <returns>A task that completes when every record has been delivered.</returns>
        /// <exception cref="ArgumentNullException"><paramref name="stream"/>, <paramref name="parser"/> or <paramref name="body"/> is <see langword="null"/>.</exception>
        /// <exception cref="ArgumentException"><paramref name="parser"/> binds columns by header name and <see cref="CsvParallelOptions.HeaderRow"/> is 0.</exception>
        /// <remarks>
        /// Carries the same contract and fallbacks as the path-based overload, and the stream restrictions
        /// of the <see cref="ICsvRecord{TSelf}"/> stream overload.
        /// </remarks>
        public static Task ForEachAsync<TModel>(
            Stream stream,
            ExcelParser<TModel> parser,
            Action<TModel> body,
            CsvParallelOptions? options = null,
            CancellationToken ct = default)
            where TModel : allows ref struct
        {
            ArgumentNullException.ThrowIfNull(stream);
            CsvParallelOptions validated = Validated(options);
            return ParallelCsvProcessor.RunAsync(stream, MappedCallback<TModel>.Unbound, BoundCallback(parser, body, validated), validated, ct);
        }

        /// <summary>
        /// Reads an in-memory CSV buffer across several threads, binds every record to a
        /// <typeparamref name="TModel"/> through <paramref name="parser"/> and folds it into a
        /// <typeparamref name="TAccumulator"/>, one instance per partition, merged in buffer order.
        /// </summary>
        /// <typeparam name="TAccumulator">The accumulator type. See <see cref="ICsvAccumulator{TSelf, TModel}"/> for the contract it must honor.</typeparam>
        /// <typeparam name="TModel">The model type.</typeparam>
        /// <param name="data">The CSV bytes. The caller keeps ownership; the buffer must not be mutated during processing.</param>
        /// <param name="parser">How columns bind to <typeparamref name="TModel"/>; created with <see cref="ExcelParser"/>. Its <see cref="ExcelParserConfig.HeaderRow"/> is ignored in favour of <see cref="CsvParallelOptions.HeaderRow"/>.</param>
        /// <param name="options">Parallelism, dialect and header options. Defaults to <see cref="CsvParallelOptions.Default"/>.</param>
        /// <param name="ct">A token to cancel processing.</param>
        /// <returns>The accumulator holding every record.</returns>
        /// <exception cref="ArgumentException"><paramref name="parser"/> binds columns by header name and <see cref="CsvParallelOptions.HeaderRow"/> is 0.</exception>
        /// <remarks>
        /// <para>
        /// The header is bound once, before any record is folded. A property that fails to parse keeps its
        /// default unless <see cref="ExcelParserConfig.ThrowOnParseFailure"/> is set; a missing
        /// <c>[ExcelRequired]</c> column or value throws <see cref="ExcelParseException"/>; empty records are skipped.
        /// </para>
        /// <para>
        /// <see cref="ExcelParseException.Row"/> is exact when the source is read sequentially and 0 when it is
        /// partitioned, since a partition does not know how many records precede it.
        /// </para>
        /// <para>Otherwise carries the same contract and fallbacks as the <see cref="ICsvRecord{TSelf}"/> overloads.</para>
        /// </remarks>
        public static Task<TAccumulator> AggregateAsync<TAccumulator, TModel>(
            ReadOnlyMemory<byte> data,
            ExcelParser<TModel> parser,
            CsvParallelOptions? options = null,
            CancellationToken ct = default)
            where TAccumulator : ICsvAccumulator<TAccumulator, TModel>, new()
            where TModel : allows ref struct
        {
            CsvParallelOptions validated = Validated(options);
            return ParallelCsvProcessor.RunAsync(data, MappedAggregation<TAccumulator, TModel>.Unbound, Bound<TAccumulator, TModel>(parser, validated), validated, ct);
        }

        /// <summary>
        /// Reads a CSV file across several threads, binds every record to a <typeparamref name="TModel"/>
        /// through <paramref name="parser"/> and folds it into a <typeparamref name="TAccumulator"/>, one instance
        /// per partition, merged in file order.
        /// </summary>
        /// <typeparam name="TAccumulator">The accumulator type. See <see cref="ICsvAccumulator{TSelf, TModel}"/> for the contract it must honor.</typeparam>
        /// <typeparam name="TModel">The model type.</typeparam>
        /// <param name="path">The path of the CSV file to read.</param>
        /// <param name="parser">How columns bind to <typeparamref name="TModel"/>; created with <see cref="ExcelParser"/>. Its <see cref="ExcelParserConfig.HeaderRow"/> is ignored in favour of <see cref="CsvParallelOptions.HeaderRow"/>.</param>
        /// <param name="options">Parallelism, dialect and header options. Defaults to <see cref="CsvParallelOptions.Default"/>.</param>
        /// <param name="ct">A token to cancel processing.</param>
        /// <returns>The accumulator holding every record.</returns>
        /// <exception cref="ArgumentException"><paramref name="parser"/> binds columns by header name and <see cref="CsvParallelOptions.HeaderRow"/> is 0.</exception>
        /// <remarks>Carries the same contract and fallbacks as the in-memory overload.</remarks>
        public static Task<TAccumulator> AggregateAsync<TAccumulator, TModel>(
            string path,
            ExcelParser<TModel> parser,
            CsvParallelOptions? options = null,
            CancellationToken ct = default)
            where TAccumulator : ICsvAccumulator<TAccumulator, TModel>, new()
            where TModel : allows ref struct
        {
            ArgumentException.ThrowIfNullOrEmpty(path);
            CsvParallelOptions validated = Validated(options);
            return ParallelCsvProcessor.RunAsync(path, MappedAggregation<TAccumulator, TModel>.Unbound, Bound<TAccumulator, TModel>(parser, validated), validated, ct);
        }

        /// <summary>
        /// Reads a CSV stream, in parallel where the stream can be partitioned, binds every record to a
        /// <typeparamref name="TModel"/> through <paramref name="parser"/> and folds it into a
        /// <typeparamref name="TAccumulator"/>, one instance per partition, merged in stream order.
        /// </summary>
        /// <typeparam name="TAccumulator">The accumulator type. See <see cref="ICsvAccumulator{TSelf, TModel}"/> for the contract it must honor.</typeparam>
        /// <typeparam name="TModel">The model type.</typeparam>
        /// <param name="stream">The CSV stream, read from its current position. The caller keeps ownership and must not read from it concurrently.</param>
        /// <param name="parser">How columns bind to <typeparamref name="TModel"/>; created with <see cref="ExcelParser"/>. Its <see cref="ExcelParserConfig.HeaderRow"/> is ignored in favour of <see cref="CsvParallelOptions.HeaderRow"/>.</param>
        /// <param name="options">Parallelism, dialect and header options. Defaults to <see cref="CsvParallelOptions.Default"/>.</param>
        /// <param name="ct">A token to cancel processing.</param>
        /// <returns>The accumulator holding every record.</returns>
        /// <exception cref="ArgumentException"><paramref name="parser"/> binds columns by header name and <see cref="CsvParallelOptions.HeaderRow"/> is 0.</exception>
        /// <remarks>Carries the same contract and fallbacks as the in-memory overload, and the stream restrictions of the <see cref="ICsvRecord{TSelf}"/> stream overload.</remarks>
        public static Task<TAccumulator> AggregateAsync<TAccumulator, TModel>(
            Stream stream,
            ExcelParser<TModel> parser,
            CsvParallelOptions? options = null,
            CancellationToken ct = default)
            where TAccumulator : ICsvAccumulator<TAccumulator, TModel>, new()
            where TModel : allows ref struct
        {
            ArgumentNullException.ThrowIfNull(stream);
            CsvParallelOptions validated = Validated(options);
            return ParallelCsvProcessor.RunAsync(stream, MappedAggregation<TAccumulator, TModel>.Unbound, Bound<TAccumulator, TModel>(parser, validated), validated, ct);
        }

        private static CsvAccumulateFactory<TAccumulator> Bound<TAccumulator, TModel>(ExcelParser<TModel> parser, CsvParallelOptions options)
            where TAccumulator : ICsvAccumulator<TAccumulator, TModel>, new()
            where TModel : allows ref struct
        {
            ArgumentNullException.ThrowIfNull(parser);
            if (options.HeaderRow == 0 && !parser.CsvInfo.IsIndexBased)
            {
                throw new ArgumentException("A parser that binds columns by header name needs CsvParallelOptions.HeaderRow of at least 1.", nameof(parser));
            }
            return MappedAggregation<TAccumulator, TModel>.Binder(parser, options.HeaderRow);
        }

        private static CsvAccumulateFactory<byte> BoundCallback<TModel>(
            ExcelParser<TModel> parser, Action<TModel> body, CsvParallelOptions options)
            where TModel : allows ref struct
        {
            ArgumentNullException.ThrowIfNull(parser);
            ArgumentNullException.ThrowIfNull(body);
            if (options.HeaderRow == 0 && !parser.CsvInfo.IsIndexBased)
            {
                throw new ArgumentException("A parser that binds columns by header name needs CsvParallelOptions.HeaderRow of at least 1.", nameof(parser));
            }
            return MappedCallback<TModel>.Binder(parser, options.HeaderRow, body);
        }

        /// <summary>
        /// Reads an in-memory CSV buffer across several threads and folds every record with the functions of
        /// <paramref name="aggregation"/>, one accumulator per partition, combined in buffer order.
        /// </summary>
        /// <typeparam name="TState">The accumulator type.</typeparam>
        /// <param name="data">The CSV bytes. The caller keeps ownership; the buffer must not be mutated during processing.</param>
        /// <param name="aggregation">The seed, accumulate and combine functions. They carry the same contract as <see cref="ICsvAccumulator{TSelf, TModel}"/>.</param>
        /// <param name="options">Parallelism, dialect and header options. Defaults to <see cref="CsvParallelOptions.Default"/>.</param>
        /// <param name="ct">A token to cancel processing.</param>
        /// <returns>The combined accumulator.</returns>
        /// <remarks>Carries the same fallbacks as the <see cref="ICsvAccumulator{TSelf, TModel}"/> overloads.</remarks>
        public static Task<TState> AggregateAsync<TState>(
            ReadOnlyMemory<byte> data,
            CsvAggregation<TState> aggregation,
            CsvParallelOptions? options = null,
            CancellationToken ct = default)
        {
            return ParallelCsvProcessor.RunAsync(data, Validated(aggregation), null, Validated(options), ct);
        }

        /// <summary>
        /// Reads a CSV file across several threads and folds every record with the functions of
        /// <paramref name="aggregation"/>, one accumulator per partition, combined in file order.
        /// </summary>
        /// <typeparam name="TState">The accumulator type.</typeparam>
        /// <param name="path">The path of the CSV file to read.</param>
        /// <param name="aggregation">The seed, accumulate and combine functions. They carry the same contract as <see cref="ICsvAccumulator{TSelf, TModel}"/>.</param>
        /// <param name="options">Parallelism, dialect and header options. Defaults to <see cref="CsvParallelOptions.Default"/>.</param>
        /// <param name="ct">A token to cancel processing.</param>
        /// <returns>The combined accumulator.</returns>
        /// <remarks>Carries the same fallbacks as the <see cref="ICsvAccumulator{TSelf, TModel}"/> overloads.</remarks>
        public static Task<TState> AggregateAsync<TState>(
            string path,
            CsvAggregation<TState> aggregation,
            CsvParallelOptions? options = null,
            CancellationToken ct = default)
        {
            ArgumentException.ThrowIfNullOrEmpty(path);
            return ParallelCsvProcessor.RunAsync(path, Validated(aggregation), null, Validated(options), ct);
        }

        /// <summary>
        /// Reads a CSV stream, in parallel where the stream can be partitioned, and folds every record with the
        /// functions of <paramref name="aggregation"/>, one accumulator per partition, combined in stream order.
        /// </summary>
        /// <typeparam name="TState">The accumulator type.</typeparam>
        /// <param name="stream">The CSV stream, read from its current position. The caller keeps ownership and must not read from it concurrently.</param>
        /// <param name="aggregation">The seed, accumulate and combine functions. They carry the same contract as <see cref="ICsvAccumulator{TSelf, TModel}"/>.</param>
        /// <param name="options">Parallelism, dialect and header options. Defaults to <see cref="CsvParallelOptions.Default"/>.</param>
        /// <param name="ct">A token to cancel processing.</param>
        /// <returns>The combined accumulator.</returns>
        /// <remarks>Carries the same fallbacks and stream restrictions as the <see cref="ICsvAccumulator{TSelf, TModel}"/> stream overload.</remarks>
        public static Task<TState> AggregateAsync<TState>(
            Stream stream,
            CsvAggregation<TState> aggregation,
            CsvParallelOptions? options = null,
            CancellationToken ct = default)
        {
            ArgumentNullException.ThrowIfNull(stream);
            return ParallelCsvProcessor.RunAsync(stream, Validated(aggregation), null, Validated(options), ct);
        }

        private static CsvAggregation<TState> Validated<TState>(CsvAggregation<TState> aggregation)
        {
            ArgumentNullException.ThrowIfNull(aggregation);
            ArgumentNullException.ThrowIfNull(aggregation.Seed, nameof(aggregation));
            ArgumentNullException.ThrowIfNull(aggregation.Accumulate, nameof(aggregation));
            ArgumentNullException.ThrowIfNull(aggregation.Combine, nameof(aggregation));
            return aggregation;
        }

        private static CsvParallelOptions Validated(CsvParallelOptions? options)
        {
            options ??= CsvParallelOptions.Default;
            ArgumentOutOfRangeException.ThrowIfNegative(options.DegreeOfParallelism, nameof(options));
            ArgumentOutOfRangeException.ThrowIfNegative(options.HeaderRow, nameof(options));
            ArgumentNullException.ThrowIfNull(options.Reader, nameof(options));
            return options;
        }

        private static CsvParallelOptions ValidatedForTypedParse(CsvParallelOptions? options)
        {
            CsvParallelOptions validated = Validated(options);
            if (validated.HeaderRow != 0)
            {
                throw new ArgumentException(
                    $"{nameof(CsvParallelOptions)}.{nameof(CsvParallelOptions.HeaderRow)} does not apply to typed parsing; set {nameof(ExcelParserConfig)}.{nameof(ExcelParserConfig.HeaderRow)} instead.",
                    nameof(options));
            }
            return validated;
        }
    }
}
