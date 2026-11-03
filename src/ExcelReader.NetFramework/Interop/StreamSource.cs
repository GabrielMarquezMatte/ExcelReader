using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;

namespace ExcelReader.NetFramework.Interop
{
    internal sealed unsafe class StreamSource
    {
        private const int BufferSize = 81920;

        private static readonly ReadFn ReadCallback = Read;
        private static readonly ReleaseFn ReleaseCallback = Release;
        private static readonly IntPtr ReadPointer = Marshal.GetFunctionPointerForDelegate(ReadCallback);
        private static readonly IntPtr ReleasePointer = Marshal.GetFunctionPointerForDelegate(ReleaseCallback);

        private static int _releaseCount;

        private readonly Stream _stream;
        private readonly byte[] _buffer = new byte[BufferSize];

        private StreamSource(Stream stream)
        {
            _stream = stream;
        }

        [UnmanagedFunctionPointer(CallingConvention.Cdecl)]
        private delegate long ReadFn(IntPtr userData, byte* buffer, long length);

        [UnmanagedFunctionPointer(CallingConvention.Cdecl)]
        private delegate void ReleaseFn(IntPtr userData);

        internal static int ReleaseCount => Volatile.Read(ref _releaseCount);

        internal static XlStream Create(Stream stream)
        {
            GCHandle handle = GCHandle.Alloc(new StreamSource(stream));
            return new XlStream
            {
                StructSize = sizeof(XlStream),
                UserData = GCHandle.ToIntPtr(handle),
                Read = ReadPointer,
                Release = ReleasePointer,
            };
        }

        [SuppressMessage("Design", "CA1031:Do not catch general exception types",
            Justification = "An exception must not cross the native callback boundary.")]
        private static long Read(IntPtr userData, byte* buffer, long length)
        {
            try
            {
                StreamSource source = (StreamSource)GCHandle.FromIntPtr(userData).Target!;
                int count = (int)Math.Min(length, source._buffer.Length);
                int read = source._stream.Read(source._buffer, 0, count);
                Marshal.Copy(source._buffer, 0, (IntPtr)buffer, read);
                return read;
            }
            catch (Exception exception)
            {
                ReportError(exception);
                return -1;
            }
        }

        private static void Release(IntPtr userData)
        {
            GCHandle.FromIntPtr(userData).Free();
            Interlocked.Increment(ref _releaseCount);
        }

        [SuppressMessage("Design", "CA1031:Do not catch general exception types",
            Justification = "An exception must not cross the native callback boundary.")]
        private static void ReportError(Exception exception)
        {
            try
            {
                byte[] message = Utf8.Bytes(exception.Message);
                fixed (byte* pointer = message)
                {
                    NativeMethods.xl_set_source_error(pointer, message.Length);
                }
            }
            catch (Exception)
            {
            }
        }
    }
}
