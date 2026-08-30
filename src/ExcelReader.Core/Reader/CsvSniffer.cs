using System.Runtime.InteropServices;
using System.Text;

namespace ExcelReader.Core.Reader
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

        /// <summary>Infers the dialect of a sample taken from the start of a delimited-text source, using <see cref="CsvSnifferOptions.Default"/>.</summary>
        /// <param name="sample">A sample of bytes from the start of the source.</param>
        /// <returns>The inferred dialect, or <see cref="CsvDialect.Default"/> when the sample does not allow a delimiter to be determined.</returns>
        public static CsvDialect Detect(ReadOnlySpan<byte> sample)
        {
            return Detect(sample, CsvSnifferOptions.Default);
        }

        /// <summary>Infers the dialect of a sample taken from the start of a delimited-text source.</summary>
        /// <param name="sample">A sample of bytes from the start of the source.</param>
        /// <param name="options">Candidate delimiters/quotes and the sample-line cap.</param>
        /// <returns>The inferred dialect, or <see cref="CsvDialect.Default"/> when the sample does not allow a delimiter to be determined.</returns>
        /// <exception cref="ArgumentNullException"><paramref name="options"/> is <see langword="null"/>.</exception>
        public static CsvDialect Detect(ReadOnlySpan<byte> sample, CsvSnifferOptions options)
        {
            ArgumentNullException.ThrowIfNull(options);
            (Encoding? encoding, bool hasBom, int bomLength) = DetectBom(sample);
            // Copied once here, not per candidate in the scoring loop below.
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
                // Delimiter/quote fall back to the default, but the detected BOM/encoding is kept.
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

        // A candidate scores only when at least one complete (newline-terminated) line was counted, and
        // the average field count exceeds 1 — a delimiter that never appears yields a field count of 1
        // on every line and must lose to one that does.
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

        // Counts fields per line, respecting quotes, up to `maxLines` complete lines. The trailing
        // segment after the last newline is never counted — it may be truncated mid-field.
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
