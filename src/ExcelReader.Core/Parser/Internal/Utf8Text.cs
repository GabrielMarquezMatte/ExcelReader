using System.Buffers;
using System.Text;

namespace ExcelReader.Core.Parser.Internal
{
    internal static class Utf8Text
    {
        internal const int StackChars = 128;

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
