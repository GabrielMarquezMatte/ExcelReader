using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace ExcelReader.Native
{
    /// <summary>
    /// One native block that doubles as it fills, so a finished column can be handed over as-is instead
    /// of copied. A <see cref="SafeHandle"/> owns the block, so a buffer dropped on an error path still
    /// frees it.
    /// </summary>
    internal sealed unsafe class GrowableNativeBuffer<T> : IDisposable where T : unmanaged
    {
        private const int InitialLength = 256;
        private Block? _block;
        private T* _data;
        private int _capacity;

        internal int Count { get; private set; }
        internal int ByteLength => checked(Count * sizeof(T));

        internal ref T Last => ref _data[Count - 1];

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        internal void Add(T value)
        {
            if (Count == _capacity)
            {
                Grow(Count + 1);
            }
            _data[Count++] = value;
        }

        internal void AddRange(ReadOnlySpan<T> values)
        {
            if (values.Length > _capacity - Count)
            {
                Grow(checked(Count + values.Length));
            }
            values.CopyTo(new Span<T>(_data + Count, values.Length));
            Count += values.Length;
        }

        internal void CopyTo(Span<byte> destination)
        {
            new ReadOnlySpan<byte>(_data, ByteLength).CopyTo(destination);
        }

        /// <summary>Hands the block to the caller, who frees it with <see cref="Marshal.FreeHGlobal"/>.</summary>
        internal IntPtr Detach()
        {
            IntPtr block = (IntPtr)_data;
            _block?.SetHandleAsInvalid();
            Reset();
            return block;
        }

        public void Dispose()
        {
            _block?.Dispose();
            Reset();
        }

        private void Reset()
        {
            _block = null;
            _data = null;
            _capacity = 0;
            Count = 0;
        }

        [MethodImpl(MethodImplOptions.NoInlining)]
        private void Grow(int minimum)
        {
            int capacity = Math.Max(minimum, Math.Max(_capacity * 2, InitialLength));
            nint bytes = checked((nint)capacity * sizeof(T));
            _block ??= new Block();
            _data = (T*)_block.Resize(bytes);
            _capacity = capacity;
        }

        private sealed class Block : SafeHandle
        {
            internal Block()
                : base(IntPtr.Zero, ownsHandle: true)
            {
            }

            public override bool IsInvalid => handle == IntPtr.Zero;

            internal IntPtr Resize(nint bytes)
            {
                SetHandle(handle == IntPtr.Zero ? Marshal.AllocHGlobal(bytes) : Marshal.ReAllocHGlobal(handle, bytes));
                return handle;
            }

            protected override bool ReleaseHandle()
            {
                Marshal.FreeHGlobal(handle);
                return true;
            }
        }
    }
}
