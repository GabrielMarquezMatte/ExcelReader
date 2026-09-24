using System.Buffers;
using System.Runtime.InteropServices;
using System.Text;

namespace ExcelReader.Core.Reader.Csv
{
    /// <summary>Infers a CSV dialect (delimiter, quote, encoding) from a sample of a delimited-text source.</summary>
    public static class CsvSniffer
    {
        private const byte Cr = (byte)'\r';
        private const byte Lf = (byte)'\n';

        private static ReadOnlySpan<byte> Utf8Bom => [0xEF, 0xBB, 0xBF];
        private static ReadOnlySpan<byte> Utf32LeBom => [0xFF, 0xFE, 0x00, 0x00];
        private static ReadOnlySpan<byte> Utf32BeBom => [0x00, 0x00, 0xFE, 0xFF];
        private static ReadOnlySpan<byte> Utf16LeBom => [0xFF, 0xFE];
        private static ReadOnlySpan<byte> Utf16BeBom => [0xFE, 0xFF];

        private const int SampleBytes = 64 * 1024;

        /// <summary>Reads a sample from the start of a seekable stream and infers its CSV dialect. The stream's position is restored before returning.</summary>
        /// <param name="stream">A seekable stream containing delimited-text data.</param>
        /// <param name="options">Candidate delimiters/quotes and the sample-line cap; <see cref="CsvSnifferOptions.Default"/> when <see langword="null"/>.</param>
        /// <exception cref="ArgumentNullException"><paramref name="stream"/> is <see langword="null"/>.</exception>
        /// <exception cref="ArgumentException"><paramref name="stream"/> does not support seeking.</exception>
        public static CsvDialect Detect(Stream stream, CsvSnifferOptions? options = null)
        {
            ArgumentNullException.ThrowIfNull(stream);
            RequireSeekableForSniff(stream);
            long start = stream.Position;
            byte[] buffer = ArrayPool<byte>.Shared.Rent(SampleBytes);
            try
            {
                int read = stream.ReadAtLeast(buffer, SampleBytes, throwOnEndOfStream: false);
                stream.Position = start;
                return Detect(buffer.AsSpan(0, read), options ?? CsvSnifferOptions.Default);
            }
            finally
            {
                ArrayPool<byte>.Shared.Return(buffer);
            }
        }

        /// <summary>Reads a sample from the start of a file and infers its CSV dialect.</summary>
        /// <param name="path">The path to the delimited-text file.</param>
        /// <param name="options">Candidate delimiters/quotes and the sample-line cap; <see cref="CsvSnifferOptions.Default"/> when <see langword="null"/>.</param>
        /// <exception cref="ArgumentNullException"><paramref name="path"/> is <see langword="null"/>.</exception>
        public static CsvDialect DetectFile(string path, CsvSnifferOptions? options = null)
        {
            ArgumentNullException.ThrowIfNull(path);
            using FileStream stream = File.OpenRead(path);
            return Detect(stream, options);
        }

        /// <summary>Asynchronously reads a sample from the start of a seekable stream and infers its CSV dialect. The stream's position is restored before returning.</summary>
        /// <param name="stream">A seekable stream containing delimited-text data.</param>
        /// <param name="options">Candidate delimiters/quotes and the sample-line cap; <see cref="CsvSnifferOptions.Default"/> when <see langword="null"/>.</param>
        /// <param name="ct">A token to cancel the read.</param>
        /// <exception cref="ArgumentNullException"><paramref name="stream"/> is <see langword="null"/>.</exception>
        /// <exception cref="ArgumentException"><paramref name="stream"/> does not support seeking.</exception>
        public static async ValueTask<CsvDialect> DetectAsync(Stream stream, CsvSnifferOptions? options = null, CancellationToken ct = default)
        {
            ArgumentNullException.ThrowIfNull(stream);
            RequireSeekableForSniff(stream);
            long start = stream.Position;
            byte[] buffer = ArrayPool<byte>.Shared.Rent(SampleBytes);
            try
            {
                int read = await stream.ReadAtLeastAsync(buffer, SampleBytes, throwOnEndOfStream: false, ct).ConfigureAwait(false);
                stream.Position = start;
                return Detect(buffer.AsSpan(0, read), options ?? CsvSnifferOptions.Default);
            }
            finally
            {
                ArrayPool<byte>.Shared.Return(buffer);
            }
        }

        /// <summary>Asynchronously reads a sample from the start of a file and infers its CSV dialect.</summary>
        /// <param name="path">The path to the delimited-text file.</param>
        /// <param name="options">Candidate delimiters/quotes and the sample-line cap; <see cref="CsvSnifferOptions.Default"/> when <see langword="null"/>.</param>
        /// <param name="ct">A token to cancel the read.</param>
        /// <exception cref="ArgumentNullException"><paramref name="path"/> is <see langword="null"/>.</exception>
        public static async ValueTask<CsvDialect> DetectFileAsync(string path, CsvSnifferOptions? options = null, CancellationToken ct = default)
        {
            ArgumentNullException.ThrowIfNull(path);
            FileStream stream = Excel.OpenAsyncFile(path);
            try
            {
                return await DetectAsync(stream, options, ct).ConfigureAwait(false);
            }
            finally
            {
                await stream.DisposeAsync().ConfigureAwait(false);
            }
        }

        private static void RequireSeekableForSniff(Stream stream)
        {
            if (!stream.CanSeek)
            {
                throw new ArgumentException(
                    "Detect requires a seekable stream so a sample can be read and the position restored. Read a sample yourself and pass it as a span for a non-seekable source.",
                    nameof(stream));
            }
        }

        /// <summary>Infers the dialect of a sample taken from the start of a delimited-text source.</summary>
        /// <param name="sample">A sample of bytes from the start of the source; only its first 64 KiB are examined.</param>
        /// <param name="options">Candidate delimiters/quotes and the sample-line cap; <see cref="CsvSnifferOptions.Default"/> when <see langword="null"/>.</param>
        /// <returns>The inferred dialect, or <see cref="CsvDialect.Default"/> when the sample does not allow a delimiter to be determined.</returns>
        public static CsvDialect Detect(ReadOnlySpan<byte> sample, CsvSnifferOptions? options = null)
        {
            options ??= CsvSnifferOptions.Default;
            if (sample.Length > SampleBytes)
            {
                sample = sample[..SampleBytes];
            }
            (Encoding? encoding, bool hasBom, int bomLength) = DetectBom(sample);
            byte[] body = sample[bomLength..].ToArray();
            byte[] delimiters = options.CandidateDelimiters;
            byte[] quotes = options.CandidateQuotes;
            int maxLines = options.MaxSampleLines;

            byte bestDelimiter = 0;
            byte bestQuote = 0;
            double bestVariance = 0;
            bool found = false;
            foreach (byte delimiter in delimiters)
            {
                foreach (byte quote in quotes)
                {
                    if (!TryScore(body, delimiter, quote, maxLines, out double variance))
                    {
                        continue;
                    }
                    if (found && variance >= bestVariance)
                    {
                        continue;
                    }
                    found = true;
                    bestVariance = variance;
                    bestDelimiter = delimiter;
                    bestQuote = quote;
                }
            }
            if (!found)
            {
                return CsvDialect.Default with { Encoding = encoding, HasByteOrderMark = hasBom };
            }
            return new CsvDialect
            {
                Delimiter = bestDelimiter,
                Quote = bestQuote,
                Encoding = encoding,
                HasByteOrderMark = hasBom,
            };
        }

        private static bool TryScore(byte[] sample, byte delimiter, byte quote, int maxLines, out double variance)
        {
            variance = 0;
            List<int> counts = CountFieldsPerLine(sample, delimiter, quote, maxLines);
            if (counts.Count == 0)
            {
                return false;
            }
            double mean = Mean(counts);
            if (mean <= 1.0)
            {
                return false;
            }
            variance = Variance(counts, mean);
            return true;
        }

        private static double Mean(List<int> counts)
        {
            long sum = 0;
            foreach (ref readonly int count in CollectionsMarshal.AsSpan(counts))
            {
                sum += count;
            }
            return sum / (double)counts.Count;
        }

        private static double Variance(List<int> counts, double mean)
        {
            double sumSquares = 0;
            foreach (ref int count in CollectionsMarshal.AsSpan(counts))
            {
                double diff = count - mean;
                sumSquares += diff * diff;
            }
            return sumSquares / counts.Count;
        }

        private static List<int> CountFieldsPerLine(byte[] sample, byte delimiter, byte quote, int maxLines)
        {
            var counts = new List<int>();
            var scanner = new CsvControlScanner(delimiter, quote);
            scanner.Reset(sample, sample.Length, 0);
            bool inQuotes = false;
            int delimiterCount = 0;
            int lineStart = 0;
            int stop = scanner.Next();
            while (stop >= 0 && counts.Count < maxLines)
            {
                stop = ProcessControlByte(sample, ref scanner, stop, delimiter, quote, ref inQuotes, ref delimiterCount, ref lineStart, counts);
            }
            return counts;
        }

        private static int ProcessControlByte(ReadOnlySpan<byte> sample, ref CsvControlScanner scanner, int stop, byte delimiter, byte quote,
            ref bool inQuotes, ref int delimiterCount, ref int lineStart, List<int> counts)
        {
            byte b = sample[stop];
            if (b == quote)
            {
                inQuotes = !inQuotes;
            }
            else if (b == delimiter)
            {
                if (!inQuotes)
                {
                    delimiterCount++;
                }
            }
            else if (!inQuotes)
            {
                counts.Add(delimiterCount + 1);
                delimiterCount = 0;
                lineStart = stop + 1;
                if (b == Cr && stop + 1 < sample.Length && sample[stop + 1] == Lf)
                {
                    lineStart++;
                    scanner.Next();
                }
            }
            return scanner.Next();
        }

        private static (Encoding? Encoding, bool HasBom, int Length) DetectBom(ReadOnlySpan<byte> sample)
        {
            if (sample.StartsWith(Utf8Bom))
            {
                return (Encoding.UTF8, true, Utf8Bom.Length);
            }
            if (sample.StartsWith(Utf32LeBom))
            {
                return (Encoding.UTF32, true, Utf32LeBom.Length);
            }
            if (sample.StartsWith(Utf32BeBom))
            {
                return (new UTF32Encoding(bigEndian: true, byteOrderMark: true), true, Utf32BeBom.Length);
            }
            if (sample.StartsWith(Utf16LeBom))
            {
                return (Encoding.Unicode, true, Utf16LeBom.Length);
            }
            if (sample.StartsWith(Utf16BeBom))
            {
                return (Encoding.BigEndianUnicode, true, Utf16BeBom.Length);
            }
            return (null, false, 0);
        }
    }
}
