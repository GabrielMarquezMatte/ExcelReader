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
            if (ex is InvalidDataException or EndOfStreamException)
            {
                return true;
            }

            if (ex is NotSupportedException)
            {
                return true;
            }

            if (ex is ExcelLimitExceededException)
            {
                return true;
            }

            if (ex is ExcelParseException)
            {
                return true;
            }

            if (ex is ExcelEncryptionException)
            {
                return true;
            }

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
                new ExcelEncryptionException(ExcelEncryptionReason.PasswordRequired, "test"),
            ];
            foreach (Exception ex in mustAccept)
            {
                if (!IsExpected(ex))
                {
                    throw new InvalidOperationException($"oracle rejects a sanctioned failure: {ex.GetType()}");
                }
            }

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
            }
        }
    }
}
