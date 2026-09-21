using System.Buffers;
using System.Runtime.CompilerServices;
using System.Text;

namespace ExcelReader.Core.Writer.Internal
{
    internal static class Utf8Chars
    {
        private const int StackChars = 256;

        [SkipLocalsInit]
        internal static void Decode<TState>(ReadOnlySpan<byte> utf8, TState state, ReadOnlySpanAction<char, TState> write)
        {
            char[]? rented = null;
            int maxChars = Encoding.UTF8.GetMaxCharCount(utf8.Length);
            Span<char> chars = maxChars <= StackChars ? stackalloc char[StackChars] : (rented = ArrayPool<char>.Shared.Rent(maxChars));
            try
            {
                write(chars[..Encoding.UTF8.GetChars(utf8, chars)], state);
            }
            finally
            {
                if (rented is not null)
                {
                    ArrayPool<char>.Shared.Return(rented);
                }
            }
        }
    }
}
