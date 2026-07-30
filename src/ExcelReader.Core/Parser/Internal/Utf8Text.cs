using System.Buffers;
using System.Text;

namespace ExcelReader.Core.Parser.Internal
{
    // Bridges Cell.Value (UTF-8 bytes) to the BCL's char-based TryParse APIs. DateTime, DateOnly,
    // TimeOnly, Guid and enum names have no IUtf8SpanParsable path on either target TFM, so the
    // bytes must be transcoded first. `stack` covers every realistic field in place; only a
    // pathologically long value reaches the pool.
    internal static class Utf8Text
    {
        // Wide enough for any ISO date/time, Guid, or enum member name.
        internal const int StackChars = 128;

        // Returns the decoded chars. When the return value came from the pool, `rented` is non-null
        // and the caller MUST pass it to Release in a finally — the returned span aliases it.
        internal static ReadOnlySpan<char> Decode(ReadOnlySpan<byte> utf8, Span<char> stack, out char[]? rented)
        {
            if (utf8.Length <= stack.Length)
            {
                rented = null;
                return stack[..Encoding.UTF8.GetChars(utf8, stack)];
            }
            rented = ArrayPool<char>.Shared.Rent(utf8.Length);
            return rented.AsSpan(0, Encoding.UTF8.GetChars(utf8, rented));
        }

        internal static void Release(char[]? rented)
        {
            if (rented is not null)
            {
                ArrayPool<char>.Shared.Return(rented);
            }
        }
    }
}
