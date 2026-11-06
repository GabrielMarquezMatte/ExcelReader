using System.Numerics;
using System.Runtime.InteropServices;

namespace ExcelReader.Core.Reader.Columnar
{
    /// <summary>
    /// Allocates and frees the native blocks a built column is made of. A test can route both through
    /// its own functions to track every block.
    /// </summary>
    internal static unsafe class ColumnBlocks
    {
        private static readonly AsyncLocal<Func<int, IntPtr>?> AllocOverrideLocal = new();
        private static readonly AsyncLocal<Action<IntPtr>?> FreeOverrideLocal = new();

        internal static Func<int, IntPtr>? AllocOverride
        {
            get
            {
                return AllocOverrideLocal.Value;
            }
            set
            {
                AllocOverrideLocal.Value = value;
            }
        }

        internal static Action<IntPtr>? FreeOverride
        {
            get
            {
                return FreeOverrideLocal.Value;
            }
            set
            {
                FreeOverrideLocal.Value = value;
            }
        }

        internal static IntPtr Alloc(int byteLength)
        {
            return AllocOverride is { } alloc ? alloc(byteLength) : Marshal.AllocHGlobal(byteLength);
        }

        internal static void Free(IntPtr block)
        {
            if (block == IntPtr.Zero)
            {
                return;
            }
            if (FreeOverride is { } free)
            {
                free(block);
                return;
            }
            Marshal.FreeHGlobal(block);
        }

        internal static IntPtr PackBitsLsbFirst(ReadOnlySpan<byte> flags)
        {
            int byteLength = Math.Max((flags.Length + 7) / 8, 1);
            IntPtr block = Marshal.AllocHGlobal(byteLength);
            Span<byte> packed = new((void*)block, byteLength);
            packed.Clear();
            for (int i = 0; i < flags.Length; i++)
            {
                if (flags[i] != 0)
                {
                    packed[i >> 3] |= (byte)(1 << (i & 7));
                }
            }
            return block;
        }

        internal static long CountUnset(IntPtr validity, long length)
        {
            if (validity == IntPtr.Zero)
            {
                return 0;
            }
            ReadOnlySpan<byte> bits = new((void*)validity, (int)((length + 7) / 8));
            ReadOnlySpan<ulong> ulongs = MemoryMarshal.Cast<byte, ulong>(bits);
            long set = 0;
            foreach (ulong packed in ulongs)
            {
                set += BitOperations.PopCount(packed);
            }
            foreach (byte packed in bits[(ulongs.Length * sizeof(ulong))..])
            {
                set += BitOperations.PopCount(packed);
            }
            return length - set;
        }
    }
}
