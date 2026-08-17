using ExcelReader.Core.Parser;
using ExcelReader.Core.Reader;

namespace ExcelReader.Fuzz
{
    /// <summary>
    /// Decides whether an exception raised while parsing arbitrary bytes is the library behaving
    /// correctly or a defect worth a crash report.
    /// </summary>
    /// <remarks>
    /// This is the whole point of the fuzz suite. Every reader is expected to reject malformed input
    /// through a small, documented set of exception types; anything outside that set means untrusted
    /// bytes reached code that assumed they were well-formed. In particular
    /// <see cref="OutOfMemoryException"/> is deliberately NOT tolerated — the reader options carry
    /// explicit resource limits (<c>MaxCellBytes</c>, <c>MaxSharedStringBytes</c>, the buffer-growth
    /// cap in <c>LimitChecks.NextBufferSize</c>) whose entire job is to convert an allocation blow-up
    /// into an <see cref="ExcelLimitExceededException"/>, so an OOM is a limit that failed to hold.
    /// </remarks>
    internal static class FuzzOracle
    {
        internal static bool IsExpected(Exception ex)
        {
            // Malformed container/records, and everything ZipArchive/DeflateStream raise for a
            // corrupt archive, all surface as InvalidDataException.
            if (ex is InvalidDataException or EndOfStreamException)
            {
                return true;
            }

            // A structurally valid file using a feature this library does not implement.
            if (ex is NotSupportedException)
            {
                return true;
            }

            // The resource limits doing exactly what they exist for.
            if (ex is ExcelLimitExceededException)
            {
                return true;
            }

            // Typed-parse failures (only reachable from the parser harnesses).
            if (ex is ExcelParseException)
            {
                return true;
            }

            // Exact types only. ArgumentOutOfRangeException derives from ArgumentException and
            // IndexOutOfRangeException is its own type — both mean an internal slice or index was
            // computed from attacker-controlled bytes without validation, which is a real defect.
            Type type = ex.GetType();
            return type == typeof(ArgumentException)
                || type == typeof(InvalidOperationException);
        }

        /// <summary>
        /// Asserts the oracle's polarity in both directions, so a run reporting zero failures means
        /// "nothing broke" rather than "the oracle accepts everything".
        /// </summary>
        /// <remarks>
        /// An over-permissive oracle is the standard silent failure of a fuzz suite: it keeps finding
        /// nothing, forever, and looks healthy while doing it. This runs before every smoke pass.
        /// </remarks>
        internal static void SelfCheck()
        {
            Exception[] mustAccept =
            [
                new InvalidDataException(),
                new NotSupportedException(),
                new ExcelLimitExceededException("MaxCellBytes", 1, 2),
                new EndOfStreamException(),
                new ArgumentException("bad option"),
            ];
            foreach (Exception ex in mustAccept)
            {
                if (!IsExpected(ex))
                {
                    throw new InvalidOperationException($"oracle rejects a sanctioned failure: {ex.GetType()}");
                }
            }

            // These are precisely the symptoms the suite exists to surface: an index or slice computed
            // from untrusted bytes, a null that "cannot" happen, and a limit that failed to hold.
            Exception[] mustReject =
            [
                new IndexOutOfRangeException(),
                new ArgumentOutOfRangeException(),
                new NullReferenceException(),
                new OutOfMemoryException(),
                new OverflowException(),
                new KeyNotFoundException(),
            ];
            foreach (Exception ex in mustReject)
            {
                if (IsExpected(ex))
                {
                    throw new InvalidOperationException($"oracle accepts a defect symptom: {ex.GetType()}");
                }
            }
        }

        /// <summary>
        /// Runs <paramref name="body"/>, swallowing only the sanctioned failures. Anything else is
        /// rethrown so the fuzzing engine records it as a crash, with the original stack preserved.
        /// </summary>
        internal static void Guard(Action body)
        {
            try
            {
                body();
            }
            catch (Exception ex) when (IsExpected(ex))
            {
                // Correct rejection of malformed input.
            }
        }
    }
}
