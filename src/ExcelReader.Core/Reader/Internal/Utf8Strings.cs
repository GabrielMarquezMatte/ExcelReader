using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.Intrinsics;
using System.Text;

namespace ExcelReader.Core.Reader.Internal
{
    internal static class Utf8Strings
    {
        private const ulong HighBits = 0x8080808080808080UL;

        private static readonly string[] SingleAscii = CreateSingleAscii();

        /// <summary>Materializes UTF-8 cell text. ASCII widens straight into the new string; Encoding.UTF8
        /// counts characters in one pass and transcodes in a second, which dominates short cell values.</summary>
        internal static string Create(ReadOnlySpan<byte> utf8)
        {
            if (BitConverter.IsLittleEndian && (uint)(utf8.Length - 1) < 15u)
            {
                return CreateShort(utf8);
            }
            if (!Ascii.IsValid(utf8))
            {
                return Encoding.UTF8.GetString(utf8);
            }
            return string.Create(utf8.Length, utf8, static (chars, ascii) => Ascii.ToUtf16(ascii, chars, out _));
        }

        // Loads the first and last word of a 1..15-byte value. The two may overlap; every load stays
        // inside the source and every store inside the new string.
        private static string CreateShort(ReadOnlySpan<byte> utf8)
        {
            int length = utf8.Length;
            ref byte source = ref MemoryMarshal.GetReference(utf8);
            ulong head;
            ulong tail;
            if (length >= 8)
            {
                head = Unsafe.ReadUnaligned<ulong>(ref source);
                tail = Unsafe.ReadUnaligned<ulong>(ref Unsafe.Add(ref source, length - 8));
                if (!IsAscii(head | tail))
                {
                    return Encoding.UTF8.GetString(utf8);
                }
                return string.Create(length, (head, tail), Widen8);
            }
            if (length >= 4)
            {
                head = Unsafe.ReadUnaligned<uint>(ref source);
                tail = Unsafe.ReadUnaligned<uint>(ref Unsafe.Add(ref source, length - 4));
                if (!IsAscii(head | tail))
                {
                    return Encoding.UTF8.GetString(utf8);
                }
                return string.Create(length, (head, tail), Widen4);
            }
            if (length >= 2)
            {
                head = Unsafe.ReadUnaligned<ushort>(ref source);
                tail = Unsafe.ReadUnaligned<ushort>(ref Unsafe.Add(ref source, length - 2));
                if (!IsAscii(head | tail))
                {
                    return Encoding.UTF8.GetString(utf8);
                }
                return string.Create(length, (head, tail), Widen2);
            }
            if (source < 0x80)
            {
                return SingleAscii[source];
            }
            return Encoding.UTF8.GetString(utf8);
        }

        private static bool IsAscii(ulong bytes)
        {
            return (bytes & HighBits) == 0;
        }

        private static void Widen8(Span<char> chars, (ulong Head, ulong Tail) words)
        {
            ref byte destination = ref Unsafe.As<char, byte>(ref MemoryMarshal.GetReference(chars));
            Vector128.WidenLower(Vector128.CreateScalarUnsafe(words.Head).AsByte()).AsByte().StoreUnsafe(ref destination);
            Vector128.WidenLower(Vector128.CreateScalarUnsafe(words.Tail).AsByte()).AsByte().StoreUnsafe(ref Unsafe.Add(ref destination, (chars.Length - 8) * 2));
        }

        private static void Widen4(Span<char> chars, (ulong Head, ulong Tail) words)
        {
            ref byte destination = ref Unsafe.As<char, byte>(ref MemoryMarshal.GetReference(chars));
            Unsafe.WriteUnaligned(ref destination, SpreadFour(words.Head));
            Unsafe.WriteUnaligned(ref Unsafe.Add(ref destination, (chars.Length - 4) * 2), SpreadFour(words.Tail));
        }

        private static void Widen2(Span<char> chars, (ulong Head, ulong Tail) words)
        {
            ref byte destination = ref Unsafe.As<char, byte>(ref MemoryMarshal.GetReference(chars));
            Unsafe.WriteUnaligned(ref destination, SpreadTwo((uint)words.Head));
            Unsafe.WriteUnaligned(ref Unsafe.Add(ref destination, (chars.Length - 2) * 2), SpreadTwo((uint)words.Tail));
        }

        private static ulong SpreadFour(ulong bytes)
        {
            bytes = (bytes | (bytes << 16)) & 0x0000FFFF0000FFFFUL;
            return (bytes | (bytes << 8)) & 0x00FF00FF00FF00FFUL;
        }

        private static uint SpreadTwo(uint bytes)
        {
            return (bytes | (bytes << 8)) & 0x00FF00FFu;
        }

        private static string[] CreateSingleAscii()
        {
            string[] table = new string[128];
            for (int c = 0; c < table.Length; c++)
            {
                table[c] = ((char)c).ToString();
            }
            return table;
        }
    }
}
