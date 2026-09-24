using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;
using System.Text;
using ExcelReader.Native.Typed;

namespace ExcelReader.Native
{
    [ExcludeFromCodeCoverage]
    internal static unsafe partial class Exports
    {
        [UnmanagedCallersOnly(EntryPoint = "xl_encrypt_package")]
        public static int EncryptPackage(byte* packagePath, int packagePathLength, byte* destinationPath, int destinationPathLength, byte* password, int passwordLength)
        {
            if (packagePath is null || packagePathLength <= 0
                || destinationPath is null || destinationPathLength <= 0
                || password is null || passwordLength <= 0)
            {
                return NativeStatus.InvalidArgument;
            }

            return NativeApi.EncryptPackage(
                new ReadOnlySpan<byte>(packagePath, packagePathLength),
                new ReadOnlySpan<byte>(destinationPath, destinationPathLength),
                new ReadOnlySpan<byte>(password, passwordLength));
        }

        [UnmanagedCallersOnly(EntryPoint = "xl_free_buffer")]
        public static void FreeBuffer(NativeBuffer* buffer)
        {
            if (buffer is null || buffer->Data == IntPtr.Zero)
            {
                return;
            }
            Marshal.FreeHGlobal(buffer->Data);
            *buffer = default;
        }

        private static void PublishBuffer(byte[]? bytes, NativeBuffer* outBuffer)
        {
            if (bytes is null || bytes.Length == 0)
            {
                return;
            }
            IntPtr data = Marshal.AllocHGlobal(bytes.Length);
            Marshal.Copy(bytes, 0, data, bytes.Length);
            outBuffer->Data = data;
            outBuffer->Length = bytes.Length;
        }

        private static bool IsValidOutBuffer(byte* buffer, int capacity, int* outLength)
        {
            if (outLength is null)
            {
                return false;
            }
            *outLength = 0;
            return capacity >= 0 && (buffer is not null || capacity == 0);
        }

        private static bool TryDecodeColumnSpecs(NativeColumnSpecRaw* specs, int specCount, out NativeColumnSpec[] decoded)
        {
            decoded = new NativeColumnSpec[specCount];
            for (int i = 0; i < specCount; i++)
            {
                NativeColumnSpecRaw raw = specs[i];
                if (!TypedApi.IsValidNameCount(raw.NameCount))
                {
                    decoded = [];
                    return false;
                }
                if (raw.NameCount > 0 && (raw.Names is null || raw.NameLens is null))
                {
                    decoded = [];
                    return false;
                }
                string[] names = new string[raw.NameCount];
                for (int n = 0; n < raw.NameCount; n++)
                {
                    if (!TypedApi.IsValidNameLength(raw.NameLens[n]))
                    {
                        decoded = [];
                        return false;
                    }
                    names[n] = Encoding.UTF8.GetString(raw.Names[n], raw.NameLens[n]);
                }
                decoded[i] = new NativeColumnSpec
                {
                    Names = names,
                    Index = raw.Index,
                    Type = raw.Type,
                    Nullable = raw.Nullable != 0,
                };
            }
            return true;
        }

        [UnmanagedCallersOnly(EntryPoint = "xl_last_error")]
        public static int LastError(byte* buffer, int capacity, int* outLength)
        {
            if (!IsValidOutBuffer(buffer, capacity, outLength))
            {
                return NativeStatus.InvalidArgument;
            }

            int status = NativeApi.LastError(new Span<byte>(buffer, capacity), out int length);
            *outLength = length;
            return status;
        }

        [UnmanagedCallersOnly(EntryPoint = "xl_last_error_ptr")]
        public static byte* LastErrorPtr(int* outLength)
        {
            if (outLength is null)
            {
                return null;
            }

            nint pointer = NativeApi.LastErrorPtr(out int length);
            *outLength = length;
            return (byte*)pointer;
        }

        [UnmanagedCallersOnly(EntryPoint = "xl_abi_version")]
        public static int AbiVersion()
        {
            return NativeStatus.AbiVersion;
        }

        internal static NativeHandle? Resolve(nint handle)
        {
            return NativeHandleTable.Resolve<NativeHandle>(handle);
        }

        internal static bool TryFree(nint handle, out NativeHandle? target)
        {
            return NativeHandleTable.TryUnregister(handle, out target);
        }
    }
}
