using System.Diagnostics.CodeAnalysis;
using System.Runtime.CompilerServices;
using System.Text;

namespace ExcelReader.Core.Reader.Internal
{
    /// <summary>
    /// Dedups materialized shared strings by table index. The slot array is created on the first
    /// <see cref="Cell.GetString"/> of a shared cell, so span-only reads never allocate it.
    /// </summary>
    internal sealed class SharedStringCache(int count)
    {
        private string?[]? _values;

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        internal bool TryGet(int index, ReadOnlySpan<byte> utf8, [NotNullWhen(true)] out string? value)
        {
            string?[]? values = Volatile.Read(ref _values);
            if (values is not null && (uint)index < (uint)values.Length)
            {
                value = Volatile.Read(ref values[index]);
                if (value is not null)
                {
                    return true;
                }
            }
            return TryAdd(index, utf8, out value);
        }

        [MethodImpl(MethodImplOptions.NoInlining)]
        private bool TryAdd(int index, ReadOnlySpan<byte> utf8, [NotNullWhen(true)] out string? value)
        {
            if ((uint)index >= (uint)count)
            {
                value = null;
                return false;
            }
            ref string? slot = ref (Volatile.Read(ref _values) ?? CreateValues())[index];
            string created = Encoding.UTF8.GetString(utf8);
            value = Interlocked.CompareExchange(ref slot, created, null) ?? created;
            return true;
        }

        private string?[] CreateValues()
        {
            string?[] created = new string?[count];
            return Interlocked.CompareExchange(ref _values, created, null) ?? created;
        }
    }
}
